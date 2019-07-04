package com.example.excel.service.serviceimp;

import com.example.excel.domain.S;
import org.springframework.stereotype.Service;
import javax.annotation.Resource;
import com.example.excel.mapper.SMapper;
import com.example.excel.service.SService;

import java.util.List;

@Service
public class SServiceImpl implements SService{

    @Resource
    private SMapper sMapper;

    @Override
    public List<S> getAllStudent() {
        List<S> list=sMapper.selectAll();
        return list;
    }
}
