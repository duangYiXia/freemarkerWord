package com.xxkj.demo.util;

import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.util.Objects;


public class ServletCommonUtil {

	public static HttpServletRequest getHttpServletRequest() {
		ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
		if (Objects.isNull(sra)) return null;
		HttpServletRequest request = sra.getRequest();
		return request;
	}

	public static HttpServletResponse getHttpServletResponse() {
		ServletRequestAttributes sra = (ServletRequestAttributes) RequestContextHolder.getRequestAttributes();
		HttpServletResponse response = sra.getResponse();
		return response;
	}
	
	public static HttpSession getHttpSession() {
		HttpServletRequest request = getHttpServletRequest();
		return request.getSession();
	}
	
	public static String getRequestContextPath() {
		HttpServletRequest request = getHttpServletRequest();
		return request.getContextPath();
	}
}
