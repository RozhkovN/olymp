package org.example.repository;

import org.example.entity.PasswordResetToken;
import org.example.entity.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.Optional;

@Repository
public interface PasswordResetTokenRepository extends JpaRepository<PasswordResetToken, Long> {
    Optional<PasswordResetToken> findByToken(String token);
    Optional<PasswordResetToken> findByUserAndUsedFalse(User user);

    @Modifying
    @Query("DELETE FROM PasswordResetToken t WHERE t.expiryDate < CURRENT_TIMESTAMP")
    void deleteExpiredTokens();

    @Modifying
    @Query("UPDATE PasswordResetToken t SET t.used = true WHERE t.user = ?1")
    void markAllTokensAsUsed(User user);

    @Modifying
    @Query("DELETE FROM PasswordResetToken t WHERE t.user = ?1")
    void deleteByUser(User user);
}