package com.home.excelexport.controllers;

import com.home.excelexport.Models.User;
import com.home.excelexport.Services.UserService;
import org.apache.commons.compress.utils.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;

@Controller
public class UserController {
    private UserService userService;

    @Autowired
    public UserController(UserService userService) {
        this.userService = userService;
    }

    @GetMapping("/")
    @ResponseBody
    public List<User> index() {
        return userService.getUsers();
    }

    @GetMapping("/excel")
    public void downloadExcel(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=customers.xlsx");
        ByteArrayInputStream stream = userService.exportExcel();
        IOUtils.copy(stream, response.getOutputStream());
    }
}
