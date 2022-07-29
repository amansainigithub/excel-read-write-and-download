package com.excel.rw.repo;

import com.excel.rw.domains.User;
import org.springframework.data.jpa.repository.JpaRepository;

public interface UserRepo  extends JpaRepository<User,Long> {
}
