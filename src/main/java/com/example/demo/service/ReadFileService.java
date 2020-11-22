package com.example.demo.service;

import java.util.List;

import org.springframework.web.multipart.MultipartFile;

import com.example.demo.User;

public interface ReadFileService {
	List<User> findAll();
	boolean saveDataFromUploadfile(MultipartFile file);

}
