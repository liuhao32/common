package com.example.demo.annotation;

import lombok.Getter;
import lombok.Setter;

import java.util.List;
import java.util.Set;

/**
 * @author YCKJ3275
 */
@Getter
@Setter
public class ImportResult<T> {

    /**
     * 解析失败的集合
     */
    private List<T> fails;
    /**
     * 解析失败的位置集合
     */
    private List<List<Integer>> failPos;
    /**
     * 解析失败的原因集合
     */
    private List<List<String>> failCause;
    /**
     * 解析成功的集合
     */
    private List<T> succs;

}
