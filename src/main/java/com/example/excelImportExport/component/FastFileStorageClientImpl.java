package com.example.excelImportExport.component;


import org.springframework.stereotype.Service;

import java.io.InputStream;

@Service
public class FastFileStorageClientImpl implements FastFileStorageClient {

    @Override
    public StorePath uploadFile(InputStream inputStream, long fileSize, String fileExtName) {
        // TODO 模拟
        StorePath storePath = new StorePath();
        storePath.setPath("");
        return storePath;
    }

    @Override
    public void deleteFile(String filePath) {

    }
}
