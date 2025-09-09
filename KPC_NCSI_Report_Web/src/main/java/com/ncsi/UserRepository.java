package com.ncsi;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;
import org.springframework.lang.NonNull;
import java.util.Optional;

@Repository
public interface UserRepository extends JpaRepository<UserLogin, String> {
    @NonNull
    Optional<UserLogin> findById(@NonNull String id);
    Optional<UserLogin> findByIdAndPassword(String id, String password);
} 