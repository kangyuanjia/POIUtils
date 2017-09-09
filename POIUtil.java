package com.pfw.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.List;
import java.util.function.Function;

import javax.servlet.http.HttpServletRequest;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.multipart.MultipartFile;

import com.pfw.entity.PageInfo;
import com.pfw.entity.ResultModel;

/**
 * POI工具类
 *
 */
public class POIUtil
{
    /**
     * page分页大小
     */
    private static final int PAGE_SIZE = 5000;

    /**
     * sheet分页大小，保证为PAGE_SIZE的倍数
     */
    private static final int SHEET_SIZE = 10000;

    /**
     * 日期String格式
     */
    public static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public static final SimpleDateFormat PATH_DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd HHmmss");

    public static final String DOWNLOAD_PATH = System.getProperty("user.home") + File.separator + "Downloads"
            + File.separator + "poi" + File.separator;

    private static final Logger log = LoggerFactory.getLogger(POIUtil.class);

    /**
     * xlsx文件后缀
     */
    public static final String[] XLSX_SUFFIXS = { ".xlsx", ".XLSX" };

    /**
     * xls文件后缀
     */
    public static final String[] XLS_SUFFIXS = { ".xls", ".XLS" };

    /**
     * 读取excel文件
     * 
     * @param headers 列头数组
     * @param file spring上传文件模型
     * @param dataFun 行级的数组组装，返回E
     * @return ResultModel
     */
    public static <E> ResultModel<E> readExcel(String[] headers, MultipartFile file, Function<Row, E> dataFun)
    {
        ResultModel<E> result = new ResultModel<E>();
        if (ArrayUtils.isEmpty(headers) || null == file || null == dataFun)
        {
            result.setMsg("参数错误！");
            return result;
        }

        boolean xlsxFlag = isSatisfySuffix(file.getOriginalFilename(), XLSX_SUFFIXS);
        if (!xlsxFlag && !isSatisfySuffix(file.getOriginalFilename(), XLS_SUFFIXS))
        {
            result.setMsg("非excel文件类型！");
            return result;
        }

        try
        {
            Sheet sheet = (xlsxFlag ? new XSSFWorkbook(file.getInputStream()) : new HSSFWorkbook(file.getInputStream()))
                    .getSheetAt(0);
            int sheetSize = sheet.getLastRowNum();
            if (sheetSize == 0)
            {
                result.setMsg("sheet页内容为空！");
                return result;
            }

            Row headerRow = sheet.getRow(0);
            for (int i = 0; i < headers.length; i++)
            {
                if (!isMatchHeader(headerRow.getCell(i), headers[i]))
                {
                    result.setMsg("模板错误，请下载最新的模板！");
                    return result;
                }
            }

            Row row = null;
            // 验证excel表是否存在空数据
            for (int i = 1; i <= sheetSize; i++)
            {
                row = sheet.getRow(i);
                for (int columnIndex = 0; columnIndex < headers.length; columnIndex++)
                {
                    if (isEmpty(row.getCell(columnIndex)))
                    {
                        result.setMsg(String.format("列[%s]-存在空数据！", headers[columnIndex]));
                        return result;
                    }
                }
            }

            for (int i = 1; i <= sheetSize; i++)
            {
                result.getList().add(dataFun.apply(sheet.getRow(i)));
            }
        }
        catch (IOException e)
        {
            log.error("IO异常..");
            result.setMsg("读取excel文件错误，请稍后再试！");
            e.printStackTrace();
        }

        result.setSuccess(true);
        return result;
    }

    /**
     * 创建Excel文件
     * 
     * @param headers 列头数组
     * @param list excel数据
     * @param dataFun 行级的数组组装，需要返回List<String>
     * @param fileName 文件名
     * @return
     */
    public static <E> File createExcel(String[] headers, List<E> list, Function<E, List<String>> dataFun,
            String fileName)
    {
        if (ArrayUtils.isEmpty(headers) || StringUtils.isEmpty(fileName) || CollectionUtils.isEmpty(list)
                || null == dataFun)
        {
            return null;
        }

        String filePath = getFilePathByFileName(fileName);
        HSSFWorkbook wb = new HSSFWorkbook();
        int size = list.size();
        int page = size / SHEET_SIZE + 1;
        int start = 0;
        int end = 0;
        for (int i = 1; i <= page; i++)
        {
            end = start + (i == page ? (size - SHEET_SIZE * (i - 1)) : SHEET_SIZE);
            List<E> sheetData = list.subList(start, end);
            Sheet sheet = wb.createSheet(String.format("第%s页", i));
            start = end;

            createSheetHeader(wb, sheet, headers);
            Row row = null;
            int sheetSize = sheetData.size();
            for (int rowNum = 0; rowNum < sheetSize; rowNum++)
            {
                List<String> rowData = dataFun.apply(sheetData.get(rowNum));
                row = sheet.createRow(rowNum + 1);
                for (int j = 0; j < rowData.size(); j++)
                {
                    row.createCell(j).setCellValue(rowData.get(j));
                }
            }
        }

        File file = null;
        FileOutputStream fos = null;
        try
        {
            file = createFile(filePath);
            fos = new FileOutputStream(file);
            wb.write(fos);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            IOUtils.closeQuietly(fos);
        }

        return file;
    }

    /**
     * 创建Excel文件
     * 
     * @param headers 列头数组
     * @param size 数据大小
     * @param queryFun 分页数据组装，传递分页信息，返回分页数据
     * @param dataFun 行级的数组组装，需要返回List &lt;String &gt;
     * @param fileName 文件名
     * @return
     */
    public static <E> void createExcelByPage(String[] headers, int size, Function<PageInfo, List<E>> queryFun,
            Function<E, List<String>> dataFun, String fileName)
    {
        if (ArrayUtils.isEmpty(headers) || StringUtils.isEmpty(fileName) || null == queryFun || null == dataFun)
        {
            return;
        }

        String filePath = getFilePathByFileName(fileName);
        HSSFWorkbook wb = new HSSFWorkbook();
        int page = size / SHEET_SIZE + 1;
        int start = 0;
        int end = 0;
        for (int i = 1; i <= page; i++)
        {
            int sheetSize = i == page ? (size - SHEET_SIZE * (i - 1)) : SHEET_SIZE;
            end = start + sheetSize;
            start = end;

            Sheet sheet = wb.createSheet(String.format("第%s页", i));
            createSheetHeader(wb, sheet, headers);
            Row row = null;
            List<E> queryData = null;
            int sheetPageCount = (sheetSize - 1) / PAGE_SIZE + 1;
            PageInfo pageInfo = new PageInfo();
            pageInfo.setPageSize(PAGE_SIZE);
            pageInfo.setRecords(size);
            for (int sheetPageIndex = 1; sheetPageIndex <= sheetPageCount; sheetPageIndex++)
            {
                // 获取分页数据
                int pageIndex = SHEET_SIZE / PAGE_SIZE * (i - 1) + sheetPageIndex;
                pageInfo.setPageIndex(pageIndex);
                queryData = queryFun.apply(pageInfo);

                // 处理分页数据
                for (int queryNum = 0; queryNum < queryData.size(); queryNum++)
                {
                    row = sheet.createRow((sheetPageIndex - 1) * PAGE_SIZE + queryNum + 1);
                    List<String> rowData = dataFun.apply(queryData.get(queryNum));
                    for (int rowNum = 0; rowNum < rowData.size(); rowNum++)
                    {
                        row.createCell(rowNum).setCellValue(rowData.get(rowNum));
                    }
                }
            }
        }

        File file = null;
        FileOutputStream fos = null;
        try
        {
            file = createFile(filePath);
            fos = new FileOutputStream(file);
            wb.write(fos);
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
        finally
        {
            IOUtils.closeQuietly(fos);
        }
    }

    /**
     * 创建excel列头
     * 
     * @param wb
     * @param sheet
     * @param headers
     */
    private static void createSheetHeader(Workbook wb, Sheet sheet, String[] headers)
    {
        if (null == sheet)
        {
            return;
        }

        Row row = sheet.createRow(0);
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

        for (int i = 0; i < headers.length; i++)
        {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(style);
        }
    }

    /**
     * 判断是否匹配excel的列头
     * 
     * @param cell
     * @param name
     * @return boolean
     */
    public static boolean isMatchHeader(Cell cell, String name)
    {
        if (null != cell && cell.toString().trim().equals(name))
        {
            return true;
        }

        return false;
    }

    /**
     * 判断Cell值是否为空
     * 
     * @param cell
     * @return boolean
     */
    public static boolean isEmpty(Cell cell)
    {
        if (null == cell || cell.toString().trim().isEmpty())
        {
            return true;
        }

        return false;
    }

    public static boolean isSatisfySuffix(String fileName, String[] suffixs)
    {
        for (String suffix : suffixs)
        {
            if (fileName.endsWith(suffix))
            {
                return true;
            }
        }

        return false;
    }

    private static File createFile(String filePath) throws IOException
    {
        File file = new File(filePath);
        if (!file.exists())
        {
            if (!file.getParentFile().exists())
            {
                file.getParentFile().mkdirs();
            }

            file.createNewFile();
        }

        return file;
    }

    public static String getFilePathByFileName(String fileName)
    {
        return DOWNLOAD_PATH + fileName;
    }

    public static String getEncodeFileName(HttpServletRequest request, String fileName)
            throws UnsupportedEncodingException
    {
        if (request.getHeader("User-Agent").indexOf("Firefox") > -1)
        {
            fileName = new String(fileName.getBytes("utf-8"), "ISO8859-1");
        }
        else
        {
            fileName = URLEncoder.encode(fileName, "utf-8");
        }

        fileName = fileName.replace("+", "%20");

        return fileName;
    }
}
