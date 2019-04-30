package me.seungwoo;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * Created by Leo.
 * User: ssw
 * Date: 2019-04-30
 * Time: 16:36
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Person {
    private String name;

    private int age;

    private String email;
}
