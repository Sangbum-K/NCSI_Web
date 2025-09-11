package com.ncsi;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;

@RestController
@RequestMapping("/api/files")
public class FileController {

    @Autowired
    private FileService fileService;

    private static final String DATA_DIR = "/home/ubuntu/Data";

    @GetMapping("/")
    public ResponseEntity<?> healthCheck() {
        return ResponseEntity.ok(Map.of("status", "FileController is running"));
    }

    @GetMapping("/base-path")
    public ResponseEntity<?> getBasePath() {
        return ResponseEntity.ok(Map.of("basePath", DATA_DIR));
    }

    @GetMapping("/total-size")
    public ResponseEntity<?> getTotalSize(@RequestParam(defaultValue = "") String path) {
        try {
            long totalSize = fileService.calculateTotalSize(path);
            return ResponseEntity.ok(Map.of("totalSize", totalSize));
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    @GetMapping("/search")
    public ResponseEntity<?> searchFiles(@RequestParam String q, @RequestParam(defaultValue = "") String path) {
        try {
            List<Map<String, Object>> results = fileService.searchFiles(q, path);
            return ResponseEntity.ok(results);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    @GetMapping("/list")
    public ResponseEntity<?> listFiles(@RequestParam(defaultValue = "") String path) {
        try {
            List<Map<String, Object>> files = fileService.listFiles(path);
            return ResponseEntity.ok(files);
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    @PostMapping("/upload")
    public ResponseEntity<?> uploadFile(
            @RequestParam("file") MultipartFile file,
            @RequestParam(defaultValue = "") String folder) {
        try {
            String result = fileService.uploadFile(file, folder);
            return ResponseEntity.ok(Map.of("message", result));
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    @GetMapping("/download")
    public ResponseEntity<Resource> downloadFile(@RequestParam String filePath) {
        try {
            Path path = Paths.get(DATA_DIR, filePath);
            Resource resource = new UrlResource(path.toUri());
            
            if (resource.exists()) {
                String contentType = Files.probeContentType(path);
                if (contentType == null) {
                    contentType = "application/octet-stream";
                }
                
                return ResponseEntity.ok()
                        .contentType(MediaType.parseMediaType(contentType))
                        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + resource.getFilename() + "\"")
                        .body(resource);
            } else {
                return ResponseEntity.notFound().build();
            }
        } catch (Exception e) {
            return ResponseEntity.badRequest().build();
        }
    }

    @DeleteMapping("/delete")
    public ResponseEntity<?> deleteFile(@RequestParam String filePath) {
        try {
            String result = fileService.deleteFile(filePath);
            return ResponseEntity.ok(Map.of("message", result));
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    @PostMapping("/create-folder")
    public ResponseEntity<?> createFolder(@RequestParam String folderPath) {
        try {
            String result = fileService.createFolder(folderPath);
            return ResponseEntity.ok(Map.of("message", result));
        } catch (Exception e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    // 간단한 파일 목록 API
    @GetMapping(value = "/api/files/list", produces = "application/json; charset=UTF-8")
    public ResponseEntity<?> getFileList(@RequestParam(defaultValue = "") String path) {
        try {
            // Content 없이 파일 목록만 반환 (빠름)
            List<Map<String, Object>> fileList = fileService.getFolderFileListOnly(path, true);
            
            return ResponseEntity.ok()
                    .header("Access-Control-Allow-Origin", "*")
                    .header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                    .header("Access-Control-Allow-Headers", "Content-Type, Authorization")
                    .header("Content-Type", "application/json; charset=UTF-8")
                    .body(fileList);
        } catch (Exception e) {
            return ResponseEntity.badRequest()
                    .header("Access-Control-Allow-Origin", "*")
                    .header("Content-Type", "application/json; charset=UTF-8")
                    .body(Map.of("error", e.getMessage()));
        }
    }

    // 폴더 구조 정보만 조회 (폴더 목록)
    @GetMapping(value = "/powerbi/folders", produces = "application/json")
    public ResponseEntity<?> getFoldersForPowerBI(@RequestParam(defaultValue = "") String parentPath) {
        try {
            List<Map<String, Object>> folders = fileService.getFoldersForPowerBI(parentPath);
            return ResponseEntity.ok()
                    .header("Access-Control-Allow-Origin", "*")
                    .header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                    .header("Access-Control-Allow-Headers", "Content-Type, Authorization")
                    .body(folders);
        } catch (Exception e) {
            return ResponseEntity.badRequest()
                    .header("Access-Control-Allow-Origin", "*")
                    .body(Map.of("error", e.getMessage()));
        }
    }

    // 개별 파일 다운로드 API (파워비아이용)
    @GetMapping(value = "/powerbi/file", produces = "application/octet-stream")
    public ResponseEntity<Resource> downloadFileForPowerBI(@RequestParam String filePath) {
        try {
            Path path = Paths.get(DATA_DIR, filePath);
            Resource resource = new UrlResource(path.toUri());
            
            if (resource.exists()) {
                String contentType = Files.probeContentType(path);
                if (contentType == null) {
                    contentType = "application/octet-stream";
                }
                
                return ResponseEntity.ok()
                        .header("Access-Control-Allow-Origin", "*")
                        .header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                        .header("Access-Control-Allow-Headers", "Content-Type, Authorization")
                        .contentType(MediaType.parseMediaType(contentType))
                        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + resource.getFilename() + "\"")
                        .body(resource);
            } else {
                return ResponseEntity.notFound()
                        .header("Access-Control-Allow-Origin", "*")
                        .build();
            }
        } catch (Exception e) {
            return ResponseEntity.badRequest()
                    .header("Access-Control-Allow-Origin", "*")
                    .build();
        }
    }

    // 폴더 내 모든 엑셀 파일의 실제 내용을 읽어서 통합하여 반환하는 API
    @GetMapping(value = "/powerbi/folder-excel-data", produces = "application/json; charset=UTF-8")
    public ResponseEntity<?> getFolderExcelDataForPowerBI(
            @RequestParam String folderPath,
            @RequestParam(defaultValue = "false") boolean includeSubfolders) {
        try {
            List<Map<String, Object>> excelData = fileService.getFolderExcelDataForPowerBI(folderPath, includeSubfolders);
            return ResponseEntity.ok()
                    .header("Access-Control-Allow-Origin", "*")
                    .header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                    .header("Access-Control-Allow-Headers", "Content-Type, Authorization")
                    .header("Content-Type", "application/json; charset=UTF-8")
                    .body(excelData);
        } catch (Exception e) {
            return ResponseEntity.badRequest()
                    .header("Access-Control-Allow-Origin", "*")
                    .header("Content-Type", "application/json; charset=UTF-8")
                    .body(Map.of("error", e.getMessage()));
        }
    }

    // CORS preflight 요청 처리
    @RequestMapping(value = "/powerbi/**", method = RequestMethod.OPTIONS)
    public ResponseEntity<?> handleCorsPreflight() {
        return ResponseEntity.ok()
                .header("Access-Control-Allow-Origin", "*")
                .header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                .header("Access-Control-Allow-Headers", "Content-Type, Authorization")
                .header("Access-Control-Max-Age", "3600")
                .build();
    }
}
