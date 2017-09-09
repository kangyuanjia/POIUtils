package com.pfw.entity;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ResultModel<E>
{

    /**
     * 状态：true，false
     * 
     */
    private boolean success;

    /**
     * 错误信息
     */
    private String msg;

    /**
     * 返回数据
     */
    private List<E> list;

    public ResultModel()
    {
        this.success = false;
        list = new ArrayList<E>();
    }

    public boolean isSuccess()
    {
        return success;
    }

    public void setSuccess(boolean success)
    {
        this.success = success;
    }

    public String getMsg()
    {
        return msg;
    }

    public void setMsg(String msg)
    {
        this.msg = msg;
    }

    public List<E> getList()
    {
        return list;
    }

    public void setList(List<E> list)
    {
        this.list = list;
    }

    public Map<String, Object> getResultInfo()
    {
        Map<String, Object> map = new HashMap<String, Object>(2);
        map.put("success", this.success);
        map.put("msg", this.msg);

        return map;
    }
}
