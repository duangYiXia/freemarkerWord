package com.xxkj.demo.util;

import com.xxkj.demo.exception.UtilException;
import freemarker.cache.ClassTemplateLoader;
import freemarker.cache.NullCacheStorage;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.TemplateExceptionHandler;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.StringWriter;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class FreemarkerUtils {
    private static final Logger log = LoggerFactory.getLogger(FreemarkerUtils.class);
    private Configuration configuration;
    private String prefix;
    private ThreadLocal<Template> currentTemplate = new ThreadLocal();
    private static final Map<String, Template> TEMPLATE_CACHE = new HashMap();

    private FreemarkerUtils(Class<?> sourceClass, String templatePath) {
        this.configuration = new Configuration(Configuration.VERSION_2_3_22);
        this.prefix = sourceClass.getName() + ":";
        this.configuration.setTemplateLoader(new ClassTemplateLoader(sourceClass, templatePath));
        this.configuration.setDefaultEncoding("UTF-8");
        this.configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        this.configuration.setCacheStorage(NullCacheStorage.INSTANCE);
    }

    public static FreemarkerUtils build(Class<?> sourceClass, String templatePath) {
        log.debug("The template path is: {}", templatePath);
        return new FreemarkerUtils(sourceClass, templatePath);
    }

    public FreemarkerUtils setConfiguration(Configuration configuration) {
        if (Objects.nonNull(configuration)) {
            this.configuration = configuration;
        }

        return this;
    }

    public FreemarkerUtils setTemplate(String templateName) {
        try {
            String cacheKey = this.prefix + templateName;
            Template template = (Template)TEMPLATE_CACHE.get(cacheKey);
            if (template == null) {
                template = this.configuration.getTemplate(templateName);
                TEMPLATE_CACHE.put(cacheKey, template);
            }

            this.currentTemplate.set(template);
            return this;
        } catch (IOException var4) {
            log.error("Get template has error! The template name is : " + templateName, var4);
            throw new UtilException("Get template has error! The template name is : " + templateName, var4);
        }
    }

    public String generate(Object model) {
        try {
            StringWriter stringWriter = new StringWriter();
            Throwable var3 = null;

            Object var6;
            try {
                BufferedWriter writer = new BufferedWriter(stringWriter);
                Throwable var5 = null;

                try {
                    this.generate(model, writer);
                    var6 = stringWriter.toString();
                } catch (Throwable var31) {
                    var6 = var31;
                    var5 = var31;
                    throw var31;
                } finally {
                    if (writer != null) {
                        if (var5 != null) {
                            try {
                                writer.close();
                            } catch (Throwable var30) {
                                var5.addSuppressed(var30);
                            }
                        } else {
                            writer.close();
                        }
                    }

                }
            } catch (Throwable var33) {
                var3 = var33;
                throw var33;
            } finally {
                if (stringWriter != null) {
                    if (var3 != null) {
                        try {
                            stringWriter.close();
                        } catch (Throwable var29) {
                            var3.addSuppressed(var29);
                        }
                    } else {
                        stringWriter.close();
                    }
                }

            }

            return (String)var6;
        } catch (IOException var35) {
            log.error("Template rendering exception!", var35);
            throw new UtilException("Template rendering exception!", var35);
        }
    }

    public void generate(Object model, File file) {
        try {
            FileOutputStream fos = new FileOutputStream(file);
            Throwable var4 = null;

            try {
                this.generate(model, (OutputStream)fos);
            } catch (Throwable var14) {
                var4 = var14;
                throw var14;
            } finally {
                if (fos != null) {
                    if (var4 != null) {
                        try {
                            fos.close();
                        } catch (Throwable var13) {
                            var4.addSuppressed(var13);
                        }
                    } else {
                        fos.close();
                    }
                }

            }

        } catch (IOException var16) {
            log.error("Create a file output stream exception!", var16);
            throw new UtilException("Create a file output stream exception!", var16);
        }
    }

    public void generate(Object model, OutputStream outputStream) {
        try {
            BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(outputStream, "utf-8"), 10240);
            Throwable var4 = null;

            try {
                this.generate(model, writer);
            } catch (Throwable var14) {
                var4 = var14;
                throw var14;
            } finally {
                if (writer != null) {
                    if (var4 != null) {
                        try {
                            writer.close();
                        } catch (Throwable var13) {
                            var4.addSuppressed(var13);
                        }
                    } else {
                        writer.close();
                    }
                }

            }

        } catch (IOException var16) {
            log.error("Create a buffered writer exception!", var16);
            throw new UtilException("Create a buffered writer exception!", var16);
        }
    }

    public void generate(Object model, BufferedWriter writer) {
        try {
            Template current = (Template)this.currentTemplate.get();
            if (current == null) {
                throw new UtilException("Current Template is null!");
            } else {
                current.process(model, writer);
                writer.flush();
            }
        } catch (IOException | TemplateException var4) {
            log.error("Template rendering exception!", var4);
            throw new UtilException("Template rendering exception!", var4);
        }
    }
}
