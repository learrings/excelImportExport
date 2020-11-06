package com.example.excelImportExport.component;


import java.io.InputStream;

/**
 * 上传文件服务器客户端
 */
public interface FastFileStorageClient {

    /**
     * 上传一般文件
     */
    StorePath uploadFile(InputStream inputStream, long fileSize, String fileExtName);

    /**
     * 删除文件.
     */
    void deleteFile(String filePath);
}
