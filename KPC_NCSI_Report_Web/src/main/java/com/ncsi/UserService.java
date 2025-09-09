package com.ncsi;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import java.util.Optional;

@Service
public class UserService {

    @Autowired
    private UserRepository userRepository;

    @Autowired
    private PasswordEncoder passwordEncoder;

    public Optional<UserLogin> authenticateAndGetUser(String id, String rawPassword) {
        return userRepository.findByIdAndPassword(id, rawPassword);
    }
}
