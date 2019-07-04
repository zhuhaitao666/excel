package com.example.excel.domain;

import java.io.Serializable;
import javax.persistence.*;
import lombok.Data;

@Data
@Table(name = "s")
public class S implements Serializable {
    @Id
    @Column(name = "id")
    private Integer id;

    @Column(name = "`name`")
    private String name;

    @Column(name = "course")
    private String course;

    @Column(name = "score")
    private Double score;

    private static final long serialVersionUID = 1L;
}