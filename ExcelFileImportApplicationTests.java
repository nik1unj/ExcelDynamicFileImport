package com.example.demo;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.repository.FileReadRepository;
import com.example.demo.service.ServiceImple;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelFileImportApplicationTests {

	@Test
	public void contextLoads() {
	}	

	 @Autowired
	  private ServiceImple service;
	  
	  @MockBean
	  private FileReadRepository repository;
	
	  @Test
	  public void findAllTest()
	  {
		  when(repository.findAll()).thenReturn(Stream.of(new User("nikunj","ramani","nikunj@tcs.com","7043559889","xls"),new User("keyur","ramani","keyur@tcs.com","87581178823","xls")).collect(Collectors.toList()));
		  assertEquals(2, service.findAll().size());
	  }
	  
	 
}
