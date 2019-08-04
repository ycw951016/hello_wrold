package name.ycw.helloworld.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

@Controller
public class HelloController {

    @GetMapping(value = "/ext")
    public String ext(){
        String s = "Hello, I wanna be a programmer!";
        return "ext";
    }
}
