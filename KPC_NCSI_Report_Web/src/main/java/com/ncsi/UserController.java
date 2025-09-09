package com.ncsi;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/api/user")
public class UserController {

    @Autowired
    private UserService userService;

    @PostMapping("/login")
    public ResponseEntity<?> login(@RequestBody Map<String, String> request) {
        String id = request.get("id");
        String password = request.get("password");
        var userOpt = userService.authenticateAndGetUser(id, password);

        Map<String, Object> response = new HashMap<>();
        if (userOpt.isPresent()) {
            response.put("success", true);
            response.put("loginLink", userOpt.get().getLoginLink());
            return ResponseEntity.ok(response);
        } else {
            response.put("success", false);
            return ResponseEntity.badRequest().body(response);
        }
    }
}
