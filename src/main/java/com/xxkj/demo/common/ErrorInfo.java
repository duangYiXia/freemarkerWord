package com.xxkj.demo.common;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by liwj0 on 2017/7/19.
 */
public enum ErrorInfo {
    TEST(90000, "测试用代码"),
    NO_RIGHT(10000, "您暂时没有权限进行相关操作"),
    TOKEN_EXP(10001, "TOKEN过期，请重新登录"),
    NO_LOGIN(10002, "该账号未登录，请重新登录"),
    PARAMS_ERROR(10003, "输入数据错误"),
    ROLE_ERROR(10004, "角色选择错误"),
    ERROR_USER_LOGIN(20001,"错误的用户名和密码"),
    CAN_NOT_DELETE_OR_UPDATE(30001,"该奖学金已经进入评定阶段，无法编辑或删除" ),
    NO_NUMBER(30002,"奖学金分配名额总数超过设置值" ),
    NO_ARRANGE(40004, "有奖学金未安排评定时间或分配名额"),
    CLOSE_GRADE_SUBMIT(40001, "该奖学金已被学院锁定进行审核，无法再提交新的学生名单"),
    CLOSE_COLLEGE_SUBMIT(40003, "该奖学金已被学校锁定进行审核，无法再提交新的学生名单"),
    FULL_NUMBER(40002, "提交的学生名单数量超过该奖学金剩余的分配名额，无法提交新的学生名单"),
    NO_REST_NUMBER(40005, "奖学金剩余名额不够申请数量"),
    NUMBER_OUT_LIMIT(30002,"申请退回的名额数量不能大于分配到的名额" ),
    ERROR_FILE_TYPE(30003, "该文件格式不被允许"),
    ERROR_FILE_UPLOAD(30005, "上传文件失败"),
    CAN_NOT_CLOSE(30004, "该奖学金仍有学生处在审核流程中，无法关闭");

    public final String desc;
    public final int code;
    // 构造方法
    private ErrorInfo(int code, String desc) {
        this.desc = desc;
        this.code = code;
    }

    private static Map<Integer,ErrorInfo> map;
    static {
        ErrorInfo[] res = values();
        map = new HashMap<>(res.length);
        for(ErrorInfo tmp: res) {
            map.put(tmp.code, tmp);
        }
    }

    public static ErrorInfo getByCode(int code){
        return map.get(code);
    }

    public static String getMsgByCode(int code){
        return map.get(code).desc;
    }
}
