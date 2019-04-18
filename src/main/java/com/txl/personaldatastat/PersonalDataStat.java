package com.txl.personaldatastat;

import com.baidu.aip.ocr.AipOcr;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.xobject.PDXObjectImage;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.core.annotation.Order;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FilenameFilter;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * @Author tangxl
 * @CreateTime 2019/4/17 17:30
 */
@Component
@Order(2)
public class PersonalDataStat implements ApplicationRunner {
    private  static Logger logger= LoggerFactory.getLogger(PersonalDataStat.class);
    @Value("${file.path}")
    private String baseFilePath;

    @Value("${appId}")
    private String appId;

    @Value("${appKey}")
    private String appKey;

    @Value("${appSecret}")
    private String appSecret;

    @Override
    public void run(ApplicationArguments args) throws Exception {
        long start=System.currentTimeMillis();
        try {
            logger.info("程序开始处理个人资料整理任务-----------");
            File basefile=new File(baseFilePath);
            if (!basefile.exists()){
                logger.info("存放文件目录不存在,任务结束");
                return;
            }
            //开始处理文件
            File excel=new File(baseFilePath+"/"+"入职资料清单_generate.xls");
            if (!excel.exists()){
                excel.createNewFile();
            }

            //获取基础目录下的所有文件夹
            File[] dirs=basefile.listFiles(new FilenameFilter() {
                @Override
                public boolean accept(File dir, String name) {
                    if (name.indexOf("入职资料清单")!=-1){
                        return false;
                    }
                    return true;
                }
            });
            WritableWorkbook book = Workbook.createWorkbook(excel);
            WritableSheet sheet=book.createSheet("sheet1",0);
            WritableCell baseCell1=new Label(0,0,"工号");
            sheet.addCell(baseCell1);
            WritableCell baseCell2=new Label(1,0,"姓名");
            sheet.addCell(baseCell2);
            WritableCell baseCell3=new Label(2,0,"身份证");
            sheet.addCell(baseCell3);
            WritableCell baseCell4=new Label(3,0,"体检报告");
            sheet.addCell(baseCell4);
            WritableCell baseCell5=new Label(4,0,"离职证明");
            sheet.addCell(baseCell5);
            WritableCell baseCell6=new Label(5,0,"毕业证");
            sheet.addCell(baseCell6);
            WritableCell baseCell7=new Label(6,0,"学位证");
            sheet.addCell(baseCell7);
            WritableCell baseCell8=new Label(7,0,"语言证书");
            sheet.addCell(baseCell8);
            WritableCell baseCell9=new Label(8,0,"职称证书");
            sheet.addCell(baseCell9);
            WritableCell baseCell10=new Label(9,0,"银行卡");
            sheet.addCell(baseCell10);
            WritableCell baseCell11=new Label(10,0,"社保卡");
            sheet.addCell(baseCell11);
            WritableCell baseCell12=new Label(11,0,"简历");
            sheet.addCell(baseCell12);
            WritableCell baseCell13=new Label(12,0,"员工登记表");
            sheet.addCell(baseCell13);
            WritableCell baseCell14=new Label(13,0,"廉洁自律书");
            sheet.addCell(baseCell14);
            WritableCell baseCell15=new Label(14,0,"三方协议");
            sheet.addCell(baseCell15);
            WritableCell baseCell16=new Label(15,0,"成绩单");
            sheet.addCell(baseCell16);
            WritableCell baseCell17=new Label(16,0,"聘用意向书");
            sheet.addCell(baseCell17);
            WritableCell baseCell18=new Label(17,0,"定薪审批表");
            sheet.addCell(baseCell18);
            WritableCell baseCell19=new Label(18,0,"转正申请表");
            sheet.addCell(baseCell19);
            WritableCell baseCell20=new Label(19,0,"异动表");
            sheet.addCell(baseCell20);
            WritableCell baseCell21=new Label(20,0,"离职资料");
            sheet.addCell(baseCell21);
            //遍历工号文件夹
            for (int i=1;i<=dirs.length;i++){
                File dir=dirs[i-1];
                if (!dir.isDirectory()){
                    continue;
                }
                WritableCell writableCell1=new Label(0,i,dir.getName());
                sheet.addCell(writableCell1);
                //获取姓名
                String name=getName(dir);
                WritableCell writableCell2=new Label(1,i,name);
                sheet.addCell(writableCell2);
                WritableCell writableCell3=CreateWritableCell(2,i,"身份证",dir);
                sheet.addCell(writableCell3);
                WritableCell writableCell4=CreateWritableCell(3,i,"体检",dir);
                sheet.addCell(writableCell4);
                WritableCell writableCell5=CreateWritableCell(4,i,"离职证明",dir);
                sheet.addCell(writableCell5);
                WritableCell writableCell6=CreateWritableCell(5,i,"毕业证",dir);
                sheet.addCell(writableCell6);
                WritableCell writableCell7=CreateWritableCell(6,i,"学位证",dir);
                sheet.addCell(writableCell7);
                WritableCell writableCell8=CreateWritableCell(7,i,"语言证书",dir);
                sheet.addCell(writableCell8);
                WritableCell writableCell9=CreateWritableCell(8,i,"职称证书",dir);
                sheet.addCell(writableCell9);
                WritableCell writableCell10=CreateWritableCell(9,i,"银行卡",dir);
                sheet.addCell(writableCell10);
                WritableCell writableCell11=CreateWritableCell(10,i,"社保卡",dir);
                sheet.addCell(writableCell11);
                WritableCell writableCell12=CreateWritableCell(11,i,"简历",dir);
                sheet.addCell(writableCell12);
                WritableCell writableCell13=CreateWritableCell(12,i,"员工登记表",dir);
                sheet.addCell(writableCell13);
                WritableCell writableCell14=CreateWritableCell(13,i,"廉洁自律书",dir);
                sheet.addCell(writableCell14);
                WritableCell writableCell15=CreateWritableCell(14,i,"三方协议",dir);
                sheet.addCell(writableCell15);
                WritableCell writableCell16=CreateWritableCell(15,i,"成绩单",dir);
                sheet.addCell(writableCell16);
                WritableCell writableCell17=CreateWritableCell(16,i,"聘用意向书",dir);
                sheet.addCell(writableCell17);
                WritableCell writableCell18=CreateWritableCell(17,i,"定薪审批表",dir);
                sheet.addCell(writableCell18);
                WritableCell writableCell19=CreateWritableCell(18,i,"转正申请表",dir);
                sheet.addCell(writableCell19);
                WritableCell writableCell20=CreateWritableCell(19,i,"异动表",dir);
                sheet.addCell(writableCell20);
                WritableCell writableCell21=CreateWritableCell(20,i,"离职资料",dir);
                sheet.addCell(writableCell21);
            }
            book.write();
            book.close();
            logger.info("程序运行完毕,生成入职整理文件成功,耗时{}ms",System.currentTimeMillis()-start);
        }catch (Exception e ){
            logger.error("生成入职资料文件失败!",e);
        }

    }

    private WritableCell CreateWritableCell(int col, int row, String identify, File dir) {
        WritableCell writableCel=null;
        File [] files=dir.listFiles();
        boolean flag=false;
        for (File file:files){
            String fileName=file.getName();
            if (fileName.indexOf(identify)!=-1){
                flag=true;
                break;
            }
        }
        if (flag){
            writableCel=new Label(col,row,"√");
        }else {
            writableCel=new Label(col,row,"×");
        }
        return writableCel;
    }

    private String  getName(File dir){
        try {
            File sfFile=null;
            for (File file:dir.listFiles()) {
                if (file.getName().indexOf("身份证") != -1) {
                    sfFile = file;
                    break;
                }
            }
            if(sfFile==null){
                return "无法获取姓名";
            }
            PDDocument document = PDDocument.load(sfFile);
            String sfImgFile="";
            List<PDPage> pages = document.getDocumentCatalog().getAllPages();
            Iterator<PDPage> iter = pages.iterator();
            int count = 0;
            while( iter.hasNext()){
                PDPage page = (PDPage)iter.next();
                PDResources resources = page.getResources();
                Map<String, PDXObjectImage> images = resources.getImages();
                if(images != null)
                {
                    Iterator<String> imageIter = images.keySet().iterator();
                    while(imageIter.hasNext())
                    {
                        count++;
                        String key = (String)imageIter.next();
                        PDXObjectImage image = (PDXObjectImage)images.get( key );
                        String name ="sf_"+System.currentTimeMillis()+""+ count;	// 图片文件名
                        sfImgFile=sfFile.getParent()+File.separator+name;
                        image.write2file(sfImgFile);// 保存图片
                    }
                }
            }
            document.close();
            //调用百度ocr接口
            AipOcr aipOcr=new AipOcr(appId,appKey,appSecret);
            // 传入可选参数调用接口
            HashMap<String, String> options = new HashMap<String, String>();
            options.put("detect_direction", "true");//检测朝向
            options.put("detect_risk", "false");//开启风险验证
            String idCardSide = "front";//front - 身份证含照片的一面 back - 身份证带国徽的一面
            JSONObject res = aipOcr.idcard(sfImgFile+".jpg",idCardSide,options);
            JSONObject wordsResult=res.getJSONObject("words_result");
            JSONObject nameResult=wordsResult.getJSONObject("姓名");
            String name=nameResult.getString("words");
            return name;
        }catch(Exception e){
            logger.error("读取pdf文件错误",e);
            return null;
        }
    }
}
