package com.demo.exceldemomvn.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

import com.demo.exceldemomvn.service.FileConversion;

@Controller
@RequestMapping("index")
public class IndexController {

	@Autowired
	FileConversion fileConversion;

	@GetMapping
	public ModelAndView index(Model model) {

		return new ModelAndView("index", "model", model);
	}

	@ResponseBody
	@GetMapping("test")
	public String test(String srcFile, String destinationFile) throws Exception {
		fileConversion.conversion(srcFile, destinationFile);
		return "hello test.";
	}

}
