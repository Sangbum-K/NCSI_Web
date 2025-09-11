package com.ncsi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ByteArrayResource;

import java.io.*;
import java.nio.file.*;
import java.util.*;

@Service
public class FileService {

    private static final String DATA_DIR = "/home/ubuntu/Data";

    public List<Map<String, Object>> listFiles(String relativePath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = relativePath.isEmpty() ? basePath : basePath.resolve(relativePath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Directory not found: " + relativePath);
        }

        List<Map<String, Object>> files = new ArrayList<>();
        
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(targetPath)) {
            for (Path path : stream) {
                Map<String, Object> fileInfo = new HashMap<>();
                fileInfo.put("name", path.getFileName().toString());
                fileInfo.put("isDirectory", Files.isDirectory(path));
                fileInfo.put("size", Files.isDirectory(path) ? 0 : Files.size(path));
                fileInfo.put("lastModified", Files.getLastModifiedTime(path).toMillis());
                
                // 상대 경로 계산
                String relativeFilePath = basePath.relativize(path).toString();
                fileInfo.put("relativePath", relativeFilePath);
                
                files.add(fileInfo);
            }
        }
        
        // 폴더를 먼저, 파일을 나중에 정렬
        files.sort((a, b) -> {
            boolean aIsDir = (Boolean) a.get("isDirectory");
            boolean bIsDir = (Boolean) b.get("isDirectory");
            
            if (aIsDir && !bIsDir) return -1;
            if (!aIsDir && bIsDir) return 1;
            
            return ((String) a.get("name")).compareToIgnoreCase((String) b.get("name"));
        });
        
        return files;
    }

    public String uploadFile(MultipartFile file, String folder) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folder.isEmpty() ? basePath : basePath.resolve(folder);
        
        // 폴더가 없으면 생성
        if (!Files.exists(targetPath)) {
            Files.createDirectories(targetPath);
        }
        
        Path filePath = targetPath.resolve(file.getOriginalFilename());
        
        // 파일이 이미 존재하면 덮어쓰기
        Files.copy(file.getInputStream(), filePath, StandardCopyOption.REPLACE_EXISTING);
        
        return "File uploaded successfully: " + file.getOriginalFilename();
    }

    public String deleteFile(String relativePath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = basePath.resolve(relativePath);
        
        if (!Files.exists(targetPath)) {
            throw new IOException("File or directory not found: " + relativePath);
        }
        
        if (Files.isDirectory(targetPath)) {
            // 폴더 삭제 (재귀적)
            deleteDirectory(targetPath);
            return "Directory deleted successfully: " + relativePath;
        } else {
            // 파일 삭제
            Files.delete(targetPath);
            return "File deleted successfully: " + relativePath;
        }
    }

    public String createFolder(String folderPath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = basePath.resolve(folderPath);
        
        if (Files.exists(targetPath)) {
            throw new IOException("Folder already exists: " + folderPath);
        }
        
        Files.createDirectories(targetPath);
        return "Folder created successfully: " + folderPath;
    }

    private void deleteDirectory(Path directory) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    deleteDirectory(path);
                } else {
                    Files.delete(path);
                }
            }
        }
        Files.delete(directory);
    }

    public long calculateTotalSize(String relativePath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = relativePath.isEmpty() ? basePath : basePath.resolve(relativePath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            return 0;
        }
        
        return calculateDirectorySize(targetPath);
    }
    
    private long calculateDirectorySize(Path directory) throws IOException {
        long totalSize = 0;
        
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    totalSize += calculateDirectorySize(path);
                } else {
                    totalSize += Files.size(path);
                }
            }
        }
        
        return totalSize;
    }

    public List<Map<String, Object>> searchFiles(String searchTerm, String relativePath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = relativePath.isEmpty() ? basePath : basePath.resolve(relativePath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            return new ArrayList<>();
        }
        
        List<Map<String, Object>> results = new ArrayList<>();
        searchInDirectory(targetPath, basePath, searchTerm.toLowerCase(), results);
        
        // 이름순으로 정렬
        results.sort((a, b) -> ((String) a.get("name")).compareToIgnoreCase((String) b.get("name")));
        
        return results;
    }
    
    private void searchInDirectory(Path directory, Path basePath, String searchTerm, List<Map<String, Object>> results) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                String fileName = path.getFileName().toString();
                
                if (fileName.toLowerCase().contains(searchTerm)) {
                    Map<String, Object> fileInfo = new HashMap<>();
                    fileInfo.put("name", fileName);
                    fileInfo.put("isDirectory", Files.isDirectory(path));
                    fileInfo.put("size", Files.isDirectory(path) ? 0 : Files.size(path));
                    fileInfo.put("lastModified", Files.getLastModifiedTime(path).toMillis());
                    
                    // 상대 경로 계산
                    String relativeFilePath = basePath.relativize(path).toString();
                    fileInfo.put("relativePath", relativeFilePath);
                    
                    results.add(fileInfo);
                }
                
                // 하위 디렉토리도 재귀적으로 검색
                if (Files.isDirectory(path)) {
                    searchInDirectory(path, basePath, searchTerm, results);
                }
            }
        }
    }

    // 파워비아이용 폴더 단위 파일 조회
    public List<Map<String, Object>> getFolderFilesForPowerBI(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> allFiles = new ArrayList<>();
        
        if (includeSubfolders) {
            // 하위 폴더까지 재귀적으로 조회
            collectAllFiles(targetPath, basePath, allFiles);
        } else {
            // 현재 폴더만 조회
            collectFilesInDirectory(targetPath, basePath, allFiles);
        }
        
        // 파워비아이 형식으로 변환
        return allFiles.stream().map(this::convertToPowerBIFormat).collect(java.util.stream.Collectors.toList());
    }

    // 폴더 목록만 조회
    public List<Map<String, Object>> getFoldersForPowerBI(String parentPath) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = parentPath.isEmpty() ? basePath : basePath.resolve(parentPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Parent folder not found: " + parentPath);
        }

        List<Map<String, Object>> folders = new ArrayList<>();
        
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(targetPath)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    Map<String, Object> folderInfo = new HashMap<>();
                    folderInfo.put("FolderName", path.getFileName().toString());
                    folderInfo.put("FolderPath", basePath.relativize(path).toString());
                    folderInfo.put("LastModified", Files.getLastModifiedTime(path).toMillis());
                    folderInfo.put("LastModifiedFormatted", new java.util.Date(Files.getLastModifiedTime(path).toMillis()).toString());
                    
                    // 폴더 내 파일 개수 계산
                    int fileCount = 0;
                    try (DirectoryStream<Path> folderStream = Files.newDirectoryStream(path)) {
                        for (Path file : folderStream) {
                            if (!Files.isDirectory(file)) {
                                fileCount++;
                            }
                        }
                    }
                    folderInfo.put("FileCount", fileCount);
                    
                    folders.add(folderInfo);
                }
            }
        }
        
        // 폴더명으로 정렬
        folders.sort((a, b) -> ((String) a.get("FolderName")).compareToIgnoreCase((String) b.get("FolderName")));
        
        return folders;
    }

    // 하위 폴더까지 모든 파일 수집
    private void collectAllFiles(Path directory, Path basePath, List<Map<String, Object>> allFiles) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                Map<String, Object> fileInfo = new HashMap<>();
                fileInfo.put("name", path.getFileName().toString());
                fileInfo.put("isDirectory", Files.isDirectory(path));
                fileInfo.put("size", Files.isDirectory(path) ? 0 : Files.size(path));
                fileInfo.put("lastModified", Files.getLastModifiedTime(path).toMillis());
                
                // 상대 경로 계산
                String relativeFilePath = basePath.relativize(path).toString();
                fileInfo.put("relativePath", relativeFilePath);
                
                // 폴더 경로 정보 추가
                String folderPath = relativeFilePath.contains("/") ? 
                    relativeFilePath.substring(0, relativeFilePath.lastIndexOf("/")) : "";
                fileInfo.put("folderPath", folderPath);
                
                allFiles.add(fileInfo);
                
                // 하위 디렉토리도 재귀적으로 처리
                if (Files.isDirectory(path)) {
                    collectAllFiles(path, basePath, allFiles);
                }
            }
        }
    }

    // 현재 폴더의 파일만 수집
    private void collectFilesInDirectory(Path directory, Path basePath, List<Map<String, Object>> files) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                Map<String, Object> fileInfo = new HashMap<>();
                fileInfo.put("name", path.getFileName().toString());
                fileInfo.put("isDirectory", Files.isDirectory(path));
                fileInfo.put("size", Files.isDirectory(path) ? 0 : Files.size(path));
                fileInfo.put("lastModified", Files.getLastModifiedTime(path).toMillis());
                
                // 상대 경로 계산
                String relativeFilePath = basePath.relativize(path).toString();
                fileInfo.put("relativePath", relativeFilePath);
                
                // 폴더 경로 정보 추가
                String folderPath = relativeFilePath.contains("/") ? 
                    relativeFilePath.substring(0, relativeFilePath.lastIndexOf("/")) : "";
                fileInfo.put("folderPath", folderPath);
                
                files.add(fileInfo);
            }
        }
    }

    // 파워비아이 형식으로 변환
    private Map<String, Object> convertToPowerBIFormat(Map<String, Object> fileInfo) {
        Map<String, Object> powerBIFile = new HashMap<>();
        powerBIFile.put("FileName", fileInfo.get("name"));
        powerBIFile.put("IsDirectory", fileInfo.get("isDirectory"));
        powerBIFile.put("FileSize", fileInfo.get("size"));
        powerBIFile.put("LastModified", fileInfo.get("lastModified"));
        powerBIFile.put("RelativePath", fileInfo.get("relativePath"));
        powerBIFile.put("FolderPath", fileInfo.get("folderPath"));
        
        // 파일 확장자 추가
        String fileName = (String) fileInfo.get("name");
        String extension = "";
        if (!(Boolean) fileInfo.get("isDirectory") && fileName.contains(".")) {
            extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
        }
        powerBIFile.put("FileExtension", extension);
        
        // 파일 크기를 읽기 쉬운 형태로 변환
        long size = (Long) fileInfo.get("size");
        powerBIFile.put("FileSizeFormatted", formatFileSize(size));
        
        // 마지막 수정일을 읽기 쉬운 형태로 변환
        long lastModified = (Long) fileInfo.get("lastModified");
        powerBIFile.put("LastModifiedFormatted", new java.util.Date(lastModified).toString());
        
        return powerBIFile;
    }

    private String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " B";
        if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
        if (bytes < 1024 * 1024 * 1024) return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
        return String.format("%.1f GB", bytes / (1024.0 * 1024.0 * 1024.0));
    }

    // 폴더 내 모든 엑셀 파일의 데이터를 한 번에 처리
    public List<Map<String, Object>> getFolderDataForPowerBI(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> allData = new ArrayList<>();
        
        if (includeSubfolders) {
            // 하위 폴더까지 재귀적으로 처리
            processAllExcelFiles(targetPath, basePath, allData);
        } else {
            // 현재 폴더만 처리
            processExcelFilesInDirectory(targetPath, basePath, allData);
        }
        
        return allData;
    }

    // 하위 폴더까지 모든 엑셀 파일 처리
    private void processAllExcelFiles(Path directory, Path basePath, List<Map<String, Object>> allData) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    // 하위 디렉토리도 재귀적으로 처리
                    processAllExcelFiles(path, basePath, allData);
                } else if (isExcelFile(path)) {
                    // 엑셀 파일 처리
                    processExcelFile(path, basePath, allData);
                }
            }
        }
    }

    // 현재 폴더의 엑셀 파일만 처리
    private void processExcelFilesInDirectory(Path directory, Path basePath, List<Map<String, Object>> allData) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    // 엑셀 파일 처리
                    processExcelFile(path, basePath, allData);
                }
            }
        }
    }

    // 엑셀 파일인지 확인
    private boolean isExcelFile(Path path) {
        String fileName = path.getFileName().toString().toLowerCase();
        return fileName.endsWith(".xlsx") || fileName.endsWith(".xls");
    }

    // 개별 엑셀 파일 처리
    private void processExcelFile(Path filePath, Path basePath, List<Map<String, Object>> allData) throws IOException {
        try {
            String relativePath = basePath.relativize(filePath).toString();
            String fileName = filePath.getFileName().toString();
            
            // 파일 정보 추가
            Map<String, Object> fileInfo = new HashMap<>();
            fileInfo.put("FileName", fileName);
            fileInfo.put("RelativePath", relativePath);
            fileInfo.put("FileSize", Files.size(filePath));
            fileInfo.put("FileSizeFormatted", formatFileSize(Files.size(filePath)));
            fileInfo.put("LastModified", Files.getLastModifiedTime(filePath).toMillis());
            fileInfo.put("LastModifiedFormatted", new java.util.Date(Files.getLastModifiedTime(filePath).toMillis()).toString());
            
            // 폴더 경로 정보
            String folderPath = relativePath.contains("/") ? 
                relativePath.substring(0, relativePath.lastIndexOf("/")) : "";
            fileInfo.put("FolderPath", folderPath);
            
            // 파일 확장자
            String extension = "";
            if (fileName.contains(".")) {
                extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
            }
            fileInfo.put("FileExtension", extension);
            
            // 다운로드 URL 추가
            fileInfo.put("DownloadUrl", "http://13.237.218.10/api/files/powerbi/file?filePath=" + relativePath);
            
            allData.add(fileInfo);
            
        } catch (Exception e) {
            // 파일 처리 중 오류가 발생해도 계속 진행
            System.err.println("Error processing file: " + filePath + " - " + e.getMessage());
        }
    }

    // 폴더 내 모든 엑셀 파일의 실제 내용을 읽어서 통합하여 반환
    public List<Map<String, Object>> getFolderExcelDataForPowerBI(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> allExcelData = new ArrayList<>();
        
        if (includeSubfolders) {
            // 하위 폴더까지 재귀적으로 처리
            processAllExcelFilesContent(targetPath, basePath, allExcelData);
        } else {
            // 현재 폴더만 처리
            processExcelFilesContentInDirectory(targetPath, basePath, allExcelData);
        }
        
        return allExcelData;
    }

    // 하위 폴더까지 모든 엑셀 파일의 내용 처리
    private void processAllExcelFilesContent(Path directory, Path basePath, List<Map<String, Object>> allData) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    // 하위 디렉토리도 재귀적으로 처리
                    processAllExcelFilesContent(path, basePath, allData);
                } else if (isExcelFile(path)) {
                    // 엑셀 파일 내용 읽기
                    readExcelFileContent(path, basePath, allData);
                }
            }
        }
    }

    // 현재 폴더의 엑셀 파일 내용만 처리
    private void processExcelFilesContentInDirectory(Path directory, Path basePath, List<Map<String, Object>> allData) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    // 엑셀 파일 내용 읽기
                    readExcelFileContent(path, basePath, allData);
                }
            }
        }
    }

    // 엑셀 파일의 실제 내용을 읽어서 데이터에 추가
    private void readExcelFileContent(Path filePath, Path basePath, List<Map<String, Object>> allData) {
        try {
            String relativePath = basePath.relativize(filePath).toString();
            String fileName = filePath.getFileName().toString();
            
            // 엑셀 파일 읽기
            try (FileInputStream fis = new FileInputStream(filePath.toFile())) {
                Workbook workbook = null;
                
                // 파일 확장자에 따라 적절한 Workbook 생성
                if (fileName.toLowerCase().endsWith(".xlsx")) {
                    workbook = new XSSFWorkbook(fis);
                } else if (fileName.toLowerCase().endsWith(".xls")) {
                    workbook = new HSSFWorkbook(fis);
                }
                
                if (workbook != null) {
                    // 모든 시트 처리
                    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                        Sheet sheet = workbook.getSheetAt(sheetIndex);
                        String sheetName = sheet.getSheetName();
                        
                        // 시트의 모든 행 처리
                        for (Row row : sheet) {
                            // 빈 행 건너뛰기
                            if (row == null) continue;
                            
                            Map<String, Object> rowData = new HashMap<>();
                            
                            // 파일 정보 추가
                            rowData.put("FileName", fileName);
                            rowData.put("RelativePath", relativePath);
                            rowData.put("SheetName", sheetName);
                            rowData.put("RowNumber", row.getRowNum() + 1);
                            
                            // 셀 데이터 추가
                            short lastCellNum = row.getLastCellNum();
                            if (lastCellNum > 0) {
                                for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
                                    Cell cell = row.getCell(cellIndex);
                                    String cellValue = getCellValueAsString(cell);
                                    rowData.put("Column" + (cellIndex + 1), cellValue);
                                }
                            }
                            
                            allData.add(rowData);
                        }
                    }
                    
                    workbook.close();
                }
            }
            
        } catch (Exception e) {
            // 파일 처리 중 오류가 발생해도 계속 진행
            System.err.println("Error reading Excel file: " + filePath + " - " + e.getMessage());
        }
    }

    // 셀 값을 문자열로 변환
    private String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    // 폴더 내 모든 Excel 파일을 하나로 합쳐서 Excel 파일로 반환
    public Resource createCombinedExcelFile(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        // 새로운 Excel 워크북 생성
        Workbook combinedWorkbook = new XSSFWorkbook();
        
        if (includeSubfolders) {
            // 하위 폴더까지 재귀적으로 처리
            combineAllExcelFiles(targetPath, basePath, combinedWorkbook);
        } else {
            // 현재 폴더만 처리
            combineExcelFilesInDirectory(targetPath, basePath, combinedWorkbook);
        }
        
        // 워크북을 바이트 배열로 변환
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            combinedWorkbook.write(outputStream);
            combinedWorkbook.close();
            
            byte[] excelBytes = outputStream.toByteArray();
            return new ByteArrayResource(excelBytes);
        }
    }

    // 하위 폴더까지 모든 Excel 파일을 하나의 워크북으로 합치기
    private void combineAllExcelFiles(Path directory, Path basePath, Workbook combinedWorkbook) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    // 하위 디렉토리도 재귀적으로 처리
                    combineAllExcelFiles(path, basePath, combinedWorkbook);
                } else if (isExcelFile(path)) {
                    // Excel 파일을 워크북에 추가
                    addExcelFileToWorkbook(path, basePath, combinedWorkbook);
                }
            }
        }
    }

    // 현재 폴더의 Excel 파일들을 하나의 워크북으로 합치기
    private void combineExcelFilesInDirectory(Path directory, Path basePath, Workbook combinedWorkbook) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    // Excel 파일을 워크북에 추가
                    addExcelFileToWorkbook(path, basePath, combinedWorkbook);
                }
            }
        }
    }

    // 개별 Excel 파일을 결합된 워크북에 추가
    private void addExcelFileToWorkbook(Path filePath, Path basePath, Workbook combinedWorkbook) {
        try {
            String relativePath = basePath.relativize(filePath).toString();
            String fileName = filePath.getFileName().toString();
            
            // 원본 Excel 파일 읽기
            try (FileInputStream fis = new FileInputStream(filePath.toFile())) {
                Workbook sourceWorkbook = null;
                
                // 파일 확장자에 따라 적절한 Workbook 생성
                if (fileName.toLowerCase().endsWith(".xlsx")) {
                    sourceWorkbook = new XSSFWorkbook(fis);
                } else if (fileName.toLowerCase().endsWith(".xls")) {
                    sourceWorkbook = new HSSFWorkbook(fis);
                }
                
                if (sourceWorkbook != null) {
                    // 원본 워크북의 모든 시트를 결합된 워크북으로 복사
                    for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
                        Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
                        String sheetName = fileName.replaceAll("\\.[^.]*$", "") + "_" + sourceSheet.getSheetName();
                        
                        // 시트명이 중복되지 않도록 처리
                        int counter = 1;
                        String originalSheetName = sheetName;
                        while (combinedWorkbook.getSheet(sheetName) != null) {
                            sheetName = originalSheetName + "_" + counter++;
                        }
                        
                        Sheet newSheet = combinedWorkbook.createSheet(sheetName);
                        copySheet(sourceSheet, newSheet);
                    }
                    
                    sourceWorkbook.close();
                }
            }
            
        } catch (Exception e) {
            // 파일 처리 중 오류가 발생해도 계속 진행
            System.err.println("Error adding Excel file to workbook: " + filePath + " - " + e.getMessage());
        }
    }

    // 시트 내용 복사
    private void copySheet(Sheet sourceSheet, Sheet targetSheet) {
        for (Row sourceRow : sourceSheet) {
            Row newRow = targetSheet.createRow(sourceRow.getRowNum());
            
            for (Cell sourceCell : sourceRow) {
                Cell newCell = newRow.createCell(sourceCell.getColumnIndex());
                
                switch (sourceCell.getCellType()) {
                    case STRING:
                        newCell.setCellValue(sourceCell.getStringCellValue());
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(sourceCell)) {
                            newCell.setCellValue(sourceCell.getDateCellValue());
                        } else {
                            newCell.setCellValue(sourceCell.getNumericCellValue());
                        }
                        break;
                    case BOOLEAN:
                        newCell.setCellValue(sourceCell.getBooleanCellValue());
                        break;
                    case FORMULA:
                        newCell.setCellFormula(sourceCell.getCellFormula());
                        break;
                    default:
                        break;
                }
                
                // 셀 스타일 복사 (선택사항)
                newCell.setCellStyle(sourceCell.getCellStyle());
            }
        }
    }

    // PowerBI Folder.Files 형식과 호환되는 파일 목록 반환
    public List<Map<String, Object>> getFolderFilesCompatible(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> fileList = new ArrayList<>();
        
        if (includeSubfolders) {
            collectFilesForPowerBI(targetPath, basePath, fileList);
        } else {
            collectFilesInDirectoryForPowerBI(targetPath, basePath, fileList);
        }
        
        return fileList;
    }

    // 하위 폴더까지 모든 Excel 파일 수집 (PowerBI 호환)
    private void collectFilesForPowerBI(Path directory, Path basePath, List<Map<String, Object>> fileList) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    collectFilesForPowerBI(path, basePath, fileList);
                } else if (isExcelFile(path)) {
                    fileList.add(createFileInfoForPowerBI(path, basePath));
                }
            }
        }
    }

    // 현재 폴더의 Excel 파일만 수집 (PowerBI 호환)
    private void collectFilesInDirectoryForPowerBI(Path directory, Path basePath, List<Map<String, Object>> fileList) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    fileList.add(createFileInfoForPowerBI(path, basePath));
                }
            }
        }
    }

    // PowerBI Folder.Files 형식에 맞는 파일 정보 생성
    private Map<String, Object> createFileInfoForPowerBI(Path filePath, Path basePath) throws IOException {
        Map<String, Object> fileInfo = new HashMap<>();
        String fileName = filePath.getFileName().toString();
        String relativePath = basePath.relativize(filePath).toString();
        
        // PowerBI Folder.Files 형식과 동일한 구조
        fileInfo.put("Name", fileName);
        fileInfo.put("Extension", fileName.contains(".") ? 
            fileName.substring(fileName.lastIndexOf(".")) : "");
        fileInfo.put("Date accessed", Files.getLastModifiedTime(filePath).toInstant().toString());
        fileInfo.put("Date modified", Files.getLastModifiedTime(filePath).toInstant().toString());
        fileInfo.put("Date created", Files.getLastModifiedTime(filePath).toInstant().toString());
        fileInfo.put("Attributes", Map.of("Directory", false, "Archive", true));
        fileInfo.put("Folder Path", basePath.relativize(filePath.getParent()).toString());
        
        // Content는 실제 Excel 파일 내용을 Base64로 인코딩
        try {
            byte[] fileContent = Files.readAllBytes(filePath);
            String base64Content = java.util.Base64.getEncoder().encodeToString(fileContent);
            fileInfo.put("Content", base64Content);
        } catch (Exception e) {
            fileInfo.put("Content", "");
        }
        
        return fileInfo;
    }

    // 제한된 수의 파일만 처리 (타임아웃 방지)
    public List<Map<String, Object>> getFolderFilesLimited(String folderPath, boolean includeSubfolders, int maxFiles) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> fileList = new ArrayList<>();
        
        if (includeSubfolders) {
            collectFilesLimited(targetPath, basePath, fileList, maxFiles);
        } else {
            collectFilesInDirectoryLimited(targetPath, basePath, fileList, maxFiles);
        }
        
        return fileList;
    }

    // 제한된 수의 파일 수집
    private void collectFilesLimited(Path directory, Path basePath, List<Map<String, Object>> fileList, int maxFiles) throws IOException {
        if (fileList.size() >= maxFiles) return;
        
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (fileList.size() >= maxFiles) break;
                
                if (Files.isDirectory(path)) {
                    collectFilesLimited(path, basePath, fileList, maxFiles);
                } else if (isExcelFile(path)) {
                    fileList.add(createFileInfoLightweight(path, basePath));
                }
            }
        }
    }

    // 현재 디렉토리에서 제한된 파일 수집
    private void collectFilesInDirectoryLimited(Path directory, Path basePath, List<Map<String, Object>> fileList, int maxFiles) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (fileList.size() >= maxFiles) break;
                
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    fileList.add(createFileInfoLightweight(path, basePath));
                }
            }
        }
    }

    // 경량화된 파일 정보 생성 (Base64 Content 포함하지만 작은 파일만)
    private Map<String, Object> createFileInfoLightweight(Path filePath, Path basePath) throws IOException {
        Map<String, Object> fileInfo = new HashMap<>();
        String fileName = filePath.getFileName().toString();
        
        fileInfo.put("Name", fileName);
        fileInfo.put("Extension", fileName.contains(".") ? 
            fileName.substring(fileName.lastIndexOf(".")) : "");
        fileInfo.put("Date modified", Files.getLastModifiedTime(filePath).toInstant().toString());
        fileInfo.put("Folder Path", basePath.relativize(filePath.getParent()).toString());
        
        // 파일 크기 체크 (10MB 이하만 Content 포함)
        long fileSize = Files.size(filePath);
        if (fileSize < 10 * 1024 * 1024) { // 10MB 이하
            try {
                byte[] fileContent = Files.readAllBytes(filePath);
                String base64Content = java.util.Base64.getEncoder().encodeToString(fileContent);
                fileInfo.put("Content", base64Content);
            } catch (Exception e) {
                fileInfo.put("Content", "");
            }
        } else {
            fileInfo.put("Content", ""); // 큰 파일은 Content 제외
        }
        
        return fileInfo;
    }

    // Content 없이 파일 목록만 반환 (빠른 조회)
    public List<Map<String, Object>> getFolderFileListOnly(String folderPath, boolean includeSubfolders) throws IOException {
        Path basePath = Paths.get(DATA_DIR);
        Path targetPath = folderPath.isEmpty() ? basePath : basePath.resolve(folderPath);
        
        if (!Files.exists(targetPath) || !Files.isDirectory(targetPath)) {
            throw new IOException("Folder not found: " + folderPath);
        }

        List<Map<String, Object>> fileList = new ArrayList<>();
        
        if (includeSubfolders) {
            collectFileListOnly(targetPath, basePath, fileList);
        } else {
            collectFileListInDirectoryOnly(targetPath, basePath, fileList);
        }
        
        return fileList;
    }

    // 파일 목록만 수집 (재귀)
    private void collectFileListOnly(Path directory, Path basePath, List<Map<String, Object>> fileList) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (Files.isDirectory(path)) {
                    collectFileListOnly(path, basePath, fileList);
                } else if (isExcelFile(path)) {
                    fileList.add(createFileInfoBasic(path, basePath));
                }
            }
        }
    }

    // 현재 디렉토리의 파일 목록만 수집
    private void collectFileListInDirectoryOnly(Path directory, Path basePath, List<Map<String, Object>> fileList) throws IOException {
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(directory)) {
            for (Path path : stream) {
                if (!Files.isDirectory(path) && isExcelFile(path)) {
                    fileList.add(createFileInfoBasic(path, basePath));
                }
            }
        }
    }

    // 기본 파일 정보만 생성 (Content 제외)
    private Map<String, Object> createFileInfoBasic(Path filePath, Path basePath) throws IOException {
        Map<String, Object> fileInfo = new HashMap<>();
        String fileName = filePath.getFileName().toString();
        String relativePath = basePath.relativize(filePath).toString();
        
        fileInfo.put("Name", fileName);
        fileInfo.put("RelativePath", relativePath);
        fileInfo.put("Extension", fileName.contains(".") ? 
            fileName.substring(fileName.lastIndexOf(".")) : "");
        fileInfo.put("Size", Files.size(filePath));
        fileInfo.put("LastModified", Files.getLastModifiedTime(filePath).toMillis());
        
        return fileInfo;
    }
}
