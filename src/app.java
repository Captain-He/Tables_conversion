import javax.swing.*;
import java.io.File;
import java.io.IOException;

public class app {
    public static String folderName= "";

    public static void main(String[] args) {
//课程目标达成评价-2019-2020-1
        folderName = args[0];
        File file = new File(folderName);
            if (file.exists()) {
                if(folderHandle(file,folderName)){
                    System.out.println(folderName+" 文件夹操作成功");
                }
            }else{
                System.out.println("文件没有检测到，请重新操作\n");
            }

        }



    /**
     * 文件夹存在，进行处理
     * @param file 文件夹
     * @return  成功的话 返回true
     */
    private static boolean folderHandle(File file,String txtName){
        boolean res = false;
        String result = "";
        File[] files = file.listFiles();
        if (null != files) {
            for (File subFile : files) {
                if (subFile.getName().endsWith("xlsx")||subFile.getName().endsWith("xls")) {
                    result += subFile.getName().substring(0,subFile.getName().lastIndexOf("."))+"\r\n";
                    result += tool.findSinglCell(subFile.getAbsolutePath(),0,1,1);
                    result += tool.findSinglCell(subFile.getAbsolutePath(),0,6,1);
                    result += tool.findSinglCell(subFile.getAbsolutePath(),0,8,1);
                    result += MainMethods.readTables(subFile.getAbsolutePath())+"\r\n";
                    System.out.println(subFile+"处理完毕");
                }
            }
            File f = new File(".");
            try {
                MainMethods.WriteToFile(result,f.getCanonicalPath()+"\\"+txtName+".txt");
                //写入成功
                res = true;
            } catch (IOException e) {
                e.printStackTrace();
            }

        }
        return res;
    }
}
