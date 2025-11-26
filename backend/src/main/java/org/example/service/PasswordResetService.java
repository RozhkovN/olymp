package org.example.service;

import org.example.dto.ResetPasswordResponse;
import org.example.entity.PasswordResetToken;
import org.example.entity.User;
import org.example.exception.UserNotFoundException;
import org.example.repository.PasswordResetTokenRepository;
import org.example.repository.UserRepository;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDateTime;
import java.util.UUID;

@Service
@Transactional
public class PasswordResetService {
    private final UserRepository userRepository;
    private final PasswordResetTokenRepository tokenRepository;
    private final BCryptPasswordEncoder passwordEncoder;
    private final EmailService emailService;

    @Value("${password.reset.token.expiry.hours:24}")
    private int tokenExpiryHours;

    public PasswordResetService(UserRepository userRepository,
                                PasswordResetTokenRepository tokenRepository,
                                EmailService emailService) {
        this.userRepository = userRepository;
        this.tokenRepository = tokenRepository;
        this.passwordEncoder = new BCryptPasswordEncoder();
        this.emailService = emailService;
    }

    public void initiatePasswordReset(String email) {
        User user = userRepository.findByEmail(email)
                .orElseThrow(() -> new UserNotFoundException("Пользователь с email " + email + " не найден"));

        // Удаляем ВСЕ токены пользователя (включая просроченные и использованные)
        tokenRepository.deleteByUser(user);

        // Создаем новый токен
        String token = UUID.randomUUID().toString();
        PasswordResetToken resetToken = new PasswordResetToken();
        resetToken.setToken(token);
        resetToken.setUser(user);
        resetToken.setExpiryDate(LocalDateTime.now().plusHours(tokenExpiryHours));
        resetToken.setUsed(false);

        tokenRepository.save(resetToken);

        // Отправляем email с ссылкой для сброса
        emailService.sendPasswordResetEmail(user.getEmail(), token);
    }

    public ResetPasswordResponse resetPassword(String token, String newPassword) {
        PasswordResetToken resetToken = tokenRepository.findByToken(token)
                .orElseThrow(() -> new RuntimeException("Неверный или просроченный токен сброса"));

        if (resetToken.isExpired()) {
            throw new RuntimeException("Срок действия токена истек");
        }

        if (resetToken.getUsed()) {
            throw new RuntimeException("Токен уже был использован");
        }

        // Обновляем пароль пользователя
        User user = resetToken.getUser();
        user.setPassword(passwordEncoder.encode(newPassword));
        userRepository.save(user);

        // Удаляем токен из базы после использования
        tokenRepository.deleteById(resetToken.getId());

        return new ResetPasswordResponse("Пароль успешно изменен", true);
    }

    public boolean validateToken(String token) {
        try {
            PasswordResetToken resetToken = tokenRepository.findByToken(token)
                    .orElseThrow(() -> new RuntimeException("Неверный токен"));

            return !resetToken.isExpired() && !resetToken.getUsed();
        } catch (Exception e) {
            return false;
        }
    }
}