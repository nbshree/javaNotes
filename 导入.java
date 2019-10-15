package com.ild.jttz.cos.exam.web.manage.user.serviecs;

import com.ild.jttz.cos.exam.model.db.dao.GbPaperUserMapMapper;
import com.ild.jttz.cos.exam.model.db.dao.GfPaperResultMapper;
import com.ild.jttz.cos.exam.model.db.dao.IdpOssAttachmentMapper;
import com.ild.jttz.cos.exam.model.db.dao.IdpSsoUserInfoMapper;
import com.ild.jttz.cos.exam.model.db.po.*;
import com.ild.jttz.cos.exam.util.ConstUtils;
import com.ild.jttz.cos.exam.web.manage.paper.services.PaperService;
import com.ild.jttz.cos.exam.web.manage.system.serviecs.DictService;
import explorer.core.reflect.ObjectConverter;
import explorer.core.util.common.StringUtil;
import explorer.data.paginator.mybatis.domain.PageBounds;
import explorer.web.session.UserSessionProvider;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import javax.annotation.Resource;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 用户信息
 */
@Service
public class UserServiceImpl implements UserService {

    @Resource
    private IdpSsoUserInfoMapper userMapper;
    private IdpOssAttachmentMapper attachmentMapper;
    @Resource
    private GbPaperUserMapMapper paperUserMapper;
    @Resource
    private DictService dictService;
    @Resource
    private PaperService paperService;
    @Resource
    private GfPaperResultMapper paperResultMapper;

    /**
     * 获取用户列表（分页）
     *
     * @param example 查询条件
     * @param pb      分页信息
     * @return
     */
    @Override
    public List<IdpSsoUserInfo> querySsoUserListPage(IdpSsoUserInfoExample example, PageBounds pb) {
        return userMapper.selectByExample(example, pb);
    }

    /**
     * 获取用户列表
     *
     * @param example 查询条件
     * @return
     */
    @Override
    public List<IdpSsoUserInfo> querySsoUserList(IdpSsoUserInfoExample example) {
        return userMapper.selectByExample(example);
    }

    /**
     * 获取用户详情
     *
     * @param id 用户id
     * @return
     */
    @Override
    public IdpSsoUserInfo loadSsoUserInfo(String id) {
        return userMapper.selectByPrimaryKey(id);
    }

    /**
     * 验证身份工号信息
     *
     * @returns
     */
    @Override
    public List<IdpSsoUserInfo> userVerification(IdpSsoUserInfoExample example) {
        return this.userMapper.selectByExample(example);
    }


    /**
     * 保存用户信息
     *
     * @param record 用户信息
     * @return
     */
    @Override
    public void saveSsoUserInfo(IdpSsoUserInfo record) {
        Date date = new Date();
        record.setModifyTime(date);
        record.setModifierId(UserSessionProvider.getUserId());
        record.setModifierName(UserSessionProvider.getUserName());
        if (StringUtil.isEmpty(record.getMobile())) {
            record.setMobile(null);
        }
        IdpSysDict entIdString = this.dictService.querySysDictListByNameOrType(record.getEntName(), "COMPANY");
        if (!StringUtil.isEmpty(entIdString.getId())) {
            record.setEntId(entIdString.getId());
        }
        if (StringUtil.isEmpty(record.getId())) {
            record.setStatus("1");
            record.setCreateTime(date);
            record.setCreatorId(UserSessionProvider.getUserId());
            record.setCreatorName(UserSessionProvider.getUserName());
            if (!StringUtil.isEmpty(record.getUserPwd())) {
                record.setUserPwd(ConstUtils.getEncryptPwd(record.getUserPwd()));
            }
            this.userMapper.insert(record);
        } else {
            IdpSsoUserInfo old = this.userMapper.selectByPrimaryKey(record.getId());
            if ("0".equals(record.getUserType()) && !"0".equals(old.getUserType())) {
                record.setUserPwd(ConstUtils.getEncryptPwd(record.getUserPwd()));
            } else {
                if (!"3".equals(record.getUserType())) {
                    if (!StringUtil.isEmpty(old.getUserPwd())) {
                        if (!old.getUserPwd().equals(record.getUserPwd())) {
                            record.setUserPwd(ConstUtils.getEncryptPwd(record.getUserPwd()));
                        }
                    }
                }
            }
            ObjectConverter.toChange(record, old);
            if (!("0".equals(record.getUserType()) && !"0".equals(old.getUserType())) && "3".equals(record.getUserType())) {
                old.setUserPwd(null);
            }
            if (StringUtil.isEmpty(record.getMobile())) {
                old.setMobile(null);
            }
            record.setModifyTime(date);
            record.setModifierId(UserSessionProvider.getUserId());
            record.setModifierName(UserSessionProvider.getUserName());
            this.userMapper.updateByPrimaryKey(old);
//            GbPaperUserMapExample gbPaperUserMapExample = new GbPaperUserMapExample();
//            GbPaperUserMapExample.Criteria criteriaPaperUserMap = gbPaperUserMapExample.createCriteria();
//            criteriaPaperUserMap.andStatusEqualTo(1);
//            criteriaPaperUserMap.andUserIdEqualTo(record.getId());
//            List<GbPaperUserMap> userList =  this.paperService.queryPaperUserGbList(gbPaperUserMapExample);
//            //修改试卷考生关联表信息
//            if (null != userList) {
//                for (GbPaperUserMap recordItem : userList) {
////                    GbPaperUserMap newRecord = new GbPaperUserMap();
//                    String recordId = recordItem.getId();
//                    ObjectConverter.toChange(old, recordItem);
//                    recordItem.setId(recordId);
//                    record.setModifyTime(date);
//                    record.setModifierId(UserSessionProvider.getUserId());
//                    record.setModifierName(UserSessionProvider.getUserName());
//                    this.paperUserMapper.updateByPrimaryKey(recordItem);
//                }
//            }
            //修改试卷考生关联表信息
            GbPaperUserMapExample gbPaperUserMapExample = new GbPaperUserMapExample();
            gbPaperUserMapExample.createCriteria().andStatusEqualTo(1).andUserIdEqualTo(record.getId());
            GbPaperUserMap newRecord = new GbPaperUserMap();
            newRecord.setUserNo(old.getUserNo());
            newRecord.setUserReal(old.getUserReal());
            newRecord.setEntId(old.getEntId());
            newRecord.setEntName(old.getEntName());
            newRecord.setPostCode(old.getPostCode());
            newRecord.setPostName(old.getPostName());
            newRecord.setEntryTime(old.getEntryTime());
            this.paperUserMapper.updateByExampleSelective(newRecord, gbPaperUserMapExample);
            //修改成绩表
            GfPaperResultExample paperResultExample = new GfPaperResultExample();
            paperResultExample.createCriteria().andStatusEqualTo(1).andUserIdEqualTo(record.getId());
            GfPaperResult paperResultRecord = new GfPaperResult();
            paperResultRecord.setUserNo(old.getUserNo());
            paperResultRecord.setUserReal(old.getUserReal());
            paperResultRecord.setEntId(old.getEntId());
            paperResultRecord.setEntName(old.getEntName());
            paperResultRecord.setPostCode(old.getPostCode());
            paperResultRecord.setPostName(old.getPostName());
            paperResultRecord.setEntryTime(old.getEntryTime());
            paperResultRecord.setUserType(old.getUserType());
            if (!StringUtil.isEmpty(record.getMobile())) {
                paperResultRecord.setMobile(old.getMobile());
            } else {
                paperResultRecord.setMobile(null);
            }
            this.paperResultMapper.updateByExampleSelective(paperResultRecord, paperResultExample);
        }
    }

    /**
     * 删除用户信息
     *
     * @param id
     * @return
     */
    @Override
    public void deleteSsoUserInfo(String id) {
        if (!StringUtil.isEmpty(id)) {
            Date now = new Date();
            IdpSsoUserInfo record = new IdpSsoUserInfo();
            record.setId(id);
            record.setStatus("0");
            record.setModifyTime(now);
            record.setModifierName(UserSessionProvider.getUserName());
            this.userMapper.updateByPrimaryKeySelective(record);
            GbPaperUserMapExample gbPaperUserMapExample = new GbPaperUserMapExample();
            GbPaperUserMapExample.Criteria criteriaPaperUserMap = gbPaperUserMapExample.createCriteria();
            criteriaPaperUserMap.andStatusEqualTo(1);
            criteriaPaperUserMap.andUserIdEqualTo(id);
            List<GbPaperUserMap> userList = this.paperService.queryPaperUserGbList(gbPaperUserMapExample);
            //修改试卷考生关联表信息
            if (null != userList) {
                for (GbPaperUserMap recordItem : userList) {
                    recordItem.setStatus(0);
                    recordItem.setModifyTime(now);
                    recordItem.setModifier(UserSessionProvider.getUserName());
                    this.paperUserMapper.updateByPrimaryKeySelective(recordItem);
                }
            }
        }
    }

    /**
     * 导入
     *
     * @param file
     * @return
     */
    @Override
    public String saveImportFileSsoUser(CommonsMultipartFile file) throws IOException, ParseException {
        HashMap<String, Integer> errorMap = new HashMap<String, Integer>();
        errorMap.put("工号未填", 0);
        errorMap.put("工号表格内重复", 0);
        errorMap.put("工号已存在", 0);
        errorMap.put("姓名未填", 0);
        errorMap.put("身份证未填", 0);
        errorMap.put("身份证表格内重复", 0);
        errorMap.put("身份证已存在", 0);
        errorMap.put("公司不存在", 0);
        errorMap.put("岗位未填", 0);
        errorMap.put("岗位不存在", 0);
        List<Integer> errorRow = new ArrayList();
        List<Integer> sucRow = new ArrayList();
        Integer count = new Integer(0);
        FileOutputStream fos = null;
        Date now = new Date();
        String successResult = new String();
        InputStream inputStream = file.getInputStream();
        List<IdpSsoUserInfo> entryList = new ArrayList<>();
        List<IdpSsoUserInfo> entryErrorList = new ArrayList<>();
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat dfo = new SimpleDateFormat("yyyyMMddHHmmss");
        try {
            XSSFWorkbook workBook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workBook.getSheetAt(0);

//            CellStyle cellStyle = workBook.createCellStyle();
//            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex()); // 前景色
//            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

            if (null != sheet) {
                for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    HashMap<String, Boolean> errorMapItem = new HashMap<String, Boolean>();
                    errorMapItem.put("工号未填", false);
                    errorMapItem.put("工号表格内重复", false);
                    errorMapItem.put("工号已存在", false);
                    errorMapItem.put("姓名未填", false);
                    errorMapItem.put("身份证未填", false);
                    errorMapItem.put("身份证表格内重复", false);
                    errorMapItem.put("身份证已存在", false);
                    errorMapItem.put("公司不存在", false);
                    errorMapItem.put("岗位未填", false);
                    errorMapItem.put("岗位不存在", false);
                    IdpSsoUserInfo info = new IdpSsoUserInfo();
                    int errorType = 0;
                    for (int i = 0; i <= 7; i++) {
                        XSSFCell cell;
                        String cellString = new String();
                        if (null != sheet.getRow(rowNum)) {
                            cell = sheet.getRow(rowNum).getCell(i);
                            switch (cell.getCellType()) {
                                case 0:      //0为CELL_TYPE_NUMERIC
                                    cellString = String.valueOf((long) cell.getNumericCellValue());
                                    break;
                                case 1:      //1为CELL_TYPE_STRING
                                case 2:      //2为CELL_TYPE_FORMULA
                                    cellString = cell.getStringCellValue();
                                    break;
                            }
                        } else {
                            cell = null;
                        }
                        switch (i) {
                            case 1://工号（必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) { //工号不能重复
                                    IdpSsoUserInfoExample example = new IdpSsoUserInfoExample();
                                    IdpSsoUserInfoExample.Criteria criteria = example.createCriteria();
                                    criteria.andStatusEqualTo("1");
                                    criteria.andUserNoEqualTo(cellString);
                                    List<IdpSsoUserInfo> IdpSsoUserInfoList = userMapper.selectByExample(example);
                                    for (int rowNumIn = 1; rowNumIn <= sheet.getLastRowNum(); rowNumIn++) {
                                        if (rowNumIn != rowNum) {
                                            XSSFCell cerNocell;
                                            String cerNocellString = new String();
                                            if (null != sheet.getRow(rowNumIn)) {
                                                cerNocell = sheet.getRow(rowNumIn).getCell(i);
                                                switch (cerNocell.getCellType()) {
                                                    case 0:      //0为CELL_TYPE_NUMERIC
                                                        cerNocellString = String.valueOf((long) cerNocell.getNumericCellValue());
                                                        break;
                                                    case 1:      //1为CELL_TYPE_STRING
                                                    case 2:      //2为CELL_TYPE_FORMULA
                                                        cerNocellString = cerNocell.getStringCellValue();
                                                        break;
                                                }
                                            } else {
                                                cerNocell = null;
                                            }
                                            if (null != cerNocell && !StringUtil.isEmpty(cerNocellString)) {
                                                if (cellString.equals(cerNocellString)) {
                                                    if (!errorRow.contains(rowNum)) {
                                                        errorRow.add(rowNum);
                                                    }
                                                    errorMapItem.put("工号表格内重复",true);
                                                    errorType += 1;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (IdpSsoUserInfoList.size() > 0) {
                                        if (!errorRow.contains(rowNum)) {
                                            errorRow.add(rowNum);
                                        }
                                        errorMapItem.put("工号已存在",true);
                                        errorType += 1;
                                    } else {
                                        info.setUserNo(cellString);
                                    }
                                } else {
                                    if (!errorRow.contains(rowNum)) {
                                        errorRow.add(rowNum);
                                    }
                                    errorMapItem.put("工号未填",true);
                                    errorType += 1;
                                }
                                break;
                            case 2://姓名（必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) {
                                    info.setUserReal(cellString);
                                } else {
                                    if (!errorRow.contains(rowNum)) {
                                        errorRow.add(rowNum);
                                    }
                                    errorMapItem.put("姓名未填",true);
                                    errorType += 1;
                                }
                                break;
                            case 3://身份证号（必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) { //身份证号不能重复
                                    IdpSsoUserInfoExample example = new IdpSsoUserInfoExample();
                                    IdpSsoUserInfoExample.Criteria criteria = example.createCriteria();
                                    criteria.andStatusEqualTo("1");
                                    criteria.andCerNoEqualTo(cellString);
                                    List<IdpSsoUserInfo> IdpSsoUserInfoList = userMapper.selectByExample(example);
                                    for (int rowNumIn = 1; rowNumIn <= sheet.getLastRowNum(); rowNumIn++) {
                                        if (rowNumIn != rowNum) {
                                            XSSFCell cerNocell;
                                            String cerNocellString = new String();
                                            if (null != sheet.getRow(rowNumIn)) {
                                                cerNocell = sheet.getRow(rowNumIn).getCell(i);
                                                switch (cerNocell.getCellType()) {
                                                    case 0:      //0为CELL_TYPE_NUMERIC
                                                        cerNocellString = String.valueOf((long) cerNocell.getNumericCellValue());
                                                        break;
                                                    case 1:      //1为CELL_TYPE_STRING
                                                    case 2:      //2为CELL_TYPE_FORMULA
                                                        cerNocellString = cerNocell.getStringCellValue();
                                                        break;
                                                }
                                            } else {
                                                cerNocell = null;
                                            }
                                            if (null != cerNocell && !StringUtil.isEmpty(cerNocellString)) {
                                                if (cellString.equals(cerNocellString)) {
                                                    if (!errorRow.contains(rowNum)) {
                                                        errorRow.add(rowNum);
                                                    }
                                                    errorType += 1;
                                                    errorMapItem.put("身份证表格内重复",true);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    if (IdpSsoUserInfoList.size() > 0) {
                                        if (!errorRow.contains(rowNum)) {
                                            errorRow.add(rowNum);
                                        }
                                        errorMapItem.put("身份证已存在",true);
                                        errorType += 1;
                                    } else {
                                        info.setCerNo(cellString);
                                    }
                                } else {
                                    if (!errorRow.contains(rowNum)) {
                                        errorRow.add(rowNum);
                                    }
                                    errorMapItem.put("身份证未填",true);
                                    errorType += 1;
                                }
                                break;
                            case 4://公司（非必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) {
                                    IdpSysDict entString = this.dictService.querySysDictListByNameOrType(cellString, "COMPANY");
                                    if (!StringUtil.isEmpty(entString.getId())) {
                                        info.setEntName(cellString);
                                        info.setEntId(entString.getId());
                                    } else {
                                        if (!errorRow.contains(rowNum)) {
                                            errorRow.add(rowNum);
                                        }
                                        errorMapItem.put("公司不存在",true);
                                        errorType += 1;
                                    }
                                }
                                break;
                            case 5://岗位（必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) {
                                    IdpSysDict postString = this.dictService.querySysDictListByNameOrType(cellString, "POST");
                                    if (!StringUtil.isEmpty(postString.getDictCode())) {
                                        info.setPostName(cellString);
                                        info.setPostCode(postString.getDictCode());
                                    } else {
                                        if (!errorRow.contains(rowNum)) {
                                            errorRow.add(rowNum);
                                        }
                                        errorMapItem.put("岗位不存在",true);
                                        errorType += 1;
                                    }
                                } else {
                                    if (!errorRow.contains(rowNum)) {
                                        errorRow.add(rowNum);
                                    }
                                    errorMapItem.put("岗位未填",true);
                                    errorType += 1;
                                }
                                break;
                            case 6://联系电话（非必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) {
                                    info.setMobile(cellString);
                                }
                                break;
                            case 7://入司时间（非必填）
                                if (null != cell && !StringUtil.isEmpty(cellString)) {
                                    DataFormatter ef = new DataFormatter();
                                    String valueAsString = ef.formatCellValue(sheet.getRow(rowNum).getCell(i));
                                    Date entryTime = df.parse(valueAsString);
                                    info.setEntryTime(entryTime);
                                }
                                break;
                        }

                    }
                    if (errorType == 0) {
                        entryList.add(info);
                        sucRow.add(rowNum);
                    } else {
                        entryErrorList.add(info);
                    }
                    for(String key:errorMapItem.keySet())
                    {
                        if(errorMapItem.get(key)){
                            errorMap.put(key,errorMap.get(key)+1);
                        }
                    }
                }
                for (int sucRowItem : sucRow) {
                    sucRowItem -= count;
                    int getLastRowNum = sheet.getLastRowNum() + 1;
                    sheet.shiftRows(sucRowItem + 1, getLastRowNum, -1);
                    count += 1;
                }
            }
            if (!entryList.isEmpty()) {
                for (IdpSsoUserInfo info : entryList) {
                    info.setId(UUID.randomUUID().toString());
                    info.setStatus("1");
                    info.setCreateTime(now);
                    info.setModifyTime(now);
                    info.setCreatorName(UserSessionProvider.getUserName());
                    info.setModifierName(UserSessionProvider.getUserName());
                    info.setRegisterSource("后台管理系统");
                    info.setUserType("3");//导入默认为非管理员
                    userMapper.insertSelective(info);
                }
            }
//            sheet.shiftRows(2,2,-1);
            StringBuilder strBuilider = new StringBuilder();
            strBuilider.append("成功导入");
            strBuilider.append(entryList.size());
            strBuilider.append("条;失败导入");
            strBuilider.append(entryErrorList.size());
            strBuilider.append("条;");
            for(String key:errorMap.keySet())
            {
                if(errorMap.get(key)>0){
                    strBuilider.append(key);
                    strBuilider.append(errorMap.get(key));
                    strBuilider.append("条;");
                }
            }
            successResult = strBuilider.toString();
//            File dest = new File("nbJttzuserFile.xlsx");
//            if (!dest.exists()) {
//                dest.createNewFile();
//            }
//            fos = new FileOutputStream(dest);
//            workBook.write(fos);// 写文件

            File uploads = new File(System.getProperty("jboss.server.data.dir"));
            if (!uploads.exists()) {
                uploads.mkdirs();
            }
            File targetFile = new File(uploads, "userError.xlsx");
            fos = new FileOutputStream(targetFile);
            workBook.write(fos);// 写文件
        } catch (Exception e) {
            throw e;
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
            if (fos != null) {
                fos.close();
            }
        }
        return successResult;
    }

    /**
     * 保存用户登陆信息
     *
     * @param record
     */
    @Override
    public void saveLoginUserInfo(IdpSsoUserInfo record) {
        if (!StringUtil.isEmpty(record.getId())) {
            userMapper.updateByPrimaryKeySelective(record);
        }
    }
}
