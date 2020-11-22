package com.example.demo;



import static org.junit.Assert.assertEquals;

import org.junit.Test;

public class Demo {
	User u = new User();
	 @Test
	    public void testfirstname() throws NoSuchFieldException, IllegalAccessException {	       
            u.setFirstName("nikunj");
	        final java.lang.reflect.Field field = u.getClass().getDeclaredField("firstName");
	        field.setAccessible(true);
	        assertEquals("Fields didn't match", field.get(u), "nikunj");
	    }
	 @Test
	    public void testlastname() throws NoSuchFieldException, IllegalAccessException {	       
       
	        u.setLastName("ramani");
	        final java.lang.reflect.Field field = u.getClass().getDeclaredField("lastName");
	        field.setAccessible(true);
	        assertEquals("Fields didn't match", field.get(u), "ramani");
	    }
	 @Test
	    public void testemail() throws NoSuchFieldException, IllegalAccessException {	       
    
	        u.setEmail("nikunj@tcs.com");
	        final java.lang.reflect.Field field = u.getClass().getDeclaredField("email");
	        field.setAccessible(true);
	        assertEquals("Fields didn't match", field.get(u), "nikunj@tcs.com");
	    }
	 @Test
	    public void testphonenumber() throws NoSuchFieldException, IllegalAccessException {	       
 
	        u.setPhoneNumber("7043559889");
	        final java.lang.reflect.Field field = u.getClass().getDeclaredField("phoneNumber");
	        field.setAccessible(true);
	        assertEquals("Fields didn't match", field.get(u), "7043559889");
	    }
}
