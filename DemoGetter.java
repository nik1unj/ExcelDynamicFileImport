package com.example.demo;

import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class DemoGetter {
	User u = new User();
	@Test
    public void getfirstname() throws NoSuchFieldException, IllegalAccessException {
        final java.lang.reflect.Field field = u.getClass().getDeclaredField("firstName");
        field.setAccessible(true);
        field.set(u, "nikunj");
        final String result = u.getFirstName();
        assertEquals("field wasn't retrieved properly", result, "nikunj");
    }
	@Test
    public void getlastname() throws NoSuchFieldException, IllegalAccessException {
        final java.lang.reflect.Field field = u.getClass().getDeclaredField("lastName");
        field.setAccessible(true);
        field.set(u, "ramani");
        final String result = u.getLastName();
        assertEquals("field wasn't retrieved properly", result, "ramani");
    }
   
	@Test
    public void getemail() throws NoSuchFieldException, IllegalAccessException {
        final java.lang.reflect.Field field = u.getClass().getDeclaredField("email");
        field.setAccessible(true);
        field.set(u, "nikunj@tcs.com");
        final String result = u.getEmail();
        assertEquals("field wasn't retrieved properly", result, "nikunj@tcs.com");
    }

	@Test
    public void getphoneNumber() throws NoSuchFieldException, IllegalAccessException {
        final java.lang.reflect.Field field = u.getClass().getDeclaredField("phoneNumber");
        field.setAccessible(true);
        field.set(u, "7043559889");
        final String result = u.getPhoneNumber();
        assertEquals("field wasn't retrieved properly", result, "7043559889");
    }


}
