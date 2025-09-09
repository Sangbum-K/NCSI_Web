package com.ncsi;

import lombok.Data;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import jakarta.persistence.Column;

@Data
@Entity
@Table(name = "User_login")
public class UserLogin {
    @Id
    private String id;
    private String password;

    @Column(name = "login_link")
    private String loginLink; // 사용자별 리포트 링크
} 