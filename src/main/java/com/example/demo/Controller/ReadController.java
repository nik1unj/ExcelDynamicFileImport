package com.example.demo.Controller;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.example.demo.User;
import com.example.demo.service.ReadFileService;

@Controller
public class ReadController {

	@Autowired
	private ReadFileService readFileService;

	@GetMapping(value = "/")
	public String home(Model model) {
		model.addAttribute("user", new User());
		List<User> users = readFileService.findAll();
		model.addAttribute("users", users);

		return "view/users";
	}

	@PostMapping(value = "/fileupload")
	public String uploadFile(@ModelAttribute User user, RedirectAttributes redirectAttributes) {
		boolean isFlag = readFileService.saveDataFromUploadfile(user.getFile());
		if (isFlag) {
			redirectAttributes.addFlashAttribute("successmessage", "File Done");

		} else {
			redirectAttributes.addFlashAttribute("errormessage", "File Not Done");

		}

		return "redirect:/";

	}
}