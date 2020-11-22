package com.example.demo.repository;

import org.springframework.data.repository.CrudRepository;
import org.springframework.stereotype.Repository;

import com.example.demo.User;

@Repository
public interface FileReadRepository extends CrudRepository<User,Long>
{

}
