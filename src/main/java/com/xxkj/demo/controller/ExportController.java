package com.xxkj.demo.controller;

import com.xxkj.demo.entity.CommonResult;
import com.xxkj.demo.util.WordUtil;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExportController {

    @GetMapping(value = "/exportRecomApprovalForm")
    public CommonResult exportRecomApprovalForm() throws Exception {
        WordUtil.main(null);

        /*if (commonResult.getStatus()) {
            return commonResult;
        } else {
            HttpServletResponse httpServletResponse = ServletCommonUtil.getHttpServletResponse();
            httpServletResponse.setHeader("error-message", URLEncoder.encode(commonResult.getMessage()));
            return new CommonResult(false, null,
                    commonResult.getCode(), commonResult.getMessage());
        }*/

        return new CommonResult() {{
            setData(null);
            setStatus(true);
            setMessage("下载成功");
            setCode(60000);
        }};
    }
}
