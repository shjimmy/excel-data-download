package cc.shjian.batch.excel;

import java.awt.GridLayout;
import java.awt.LayoutManager;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import sun.rmi.runtime.Log;

public class JFrameDownloadPhoneImg extends JFrame {
	
	JPanel panel;
	
	LinkedList<String> logList = new LinkedList<String>();
	
	JButton buttonFile,buttonOut;
	JButton buttonRead;
	JButton buttonDownload,buttonStop;
	JTextArea textareaLog;
	
	JTextField textfieldFile,textfieldOut;
	JTextField textfieldDirName;
	//JTextField textfieldColIndex;
	
	File selectFile;
	File selectOutDir;
	List<DownloadData> dataList;
	int downloadCount;
	
	int state = 0;//0=δ������ֹͣ��1=�����У�2=��ͣ
	
	JFrameDownloadPhoneImg(){
		
		//���óߴ�
		setSize(600, 620);
		
		//����Ļ����
		setLocationRelativeTo(null);
		
		//�̶������С
		setResizable(false);
		
		//�ر�ʱ�Ĳ���
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
        initPanel();//��ʼ�����
        
        add(this.panel);
        
        addLog("��ѡ���ļ�");
	}

	private void initPanel() {
		this.panel = new JPanel();
		this.panel.setLayout(null);
		
		JLabel labelFile = new JLabel("�����ļ�(Excel):");  
		textfieldFile = new JTextField(10);
		buttonFile = new JButton("ѡ��");
		 
		labelFile.setBounds(10,20,120,30);
		textfieldFile.setBounds(120,20,400,30);
		buttonFile.setBounds(520,20,50, 30);
		textfieldFile.setEditable(false);
		buttonFile.addActionListener(actionListener);
        this.panel.add(labelFile);
        this.panel.add(textfieldFile);
        this.panel.add(buttonFile);
        
		JLabel labelOut = new JLabel("�����ļ���:");  
		textfieldOut = new JTextField(10);
		buttonOut = new JButton("ѡ��");
		 
		labelOut.setBounds(10,60,120,30);
		textfieldOut.setBounds(120,60,400,30);
		buttonOut.setBounds(520,60,50, 30);
		textfieldOut.setEditable(false);
		buttonOut.addActionListener(actionListener);
        this.panel.add(labelOut);
        this.panel.add(textfieldOut);
        this.panel.add(buttonOut);
        
        buttonRead = new JButton("��ȡ�ļ�");
        buttonRead.setBounds(10,100, 570, 45);
        buttonRead.addActionListener(actionListener);
        this.panel.add(buttonRead);
        
        
        JLabel labelDirName = new JLabel("�ļ��������õ�");  
        textfieldDirName = new JTextField(10);
		JLabel labelDirName2 = new JLabel("������");  
		
		labelDirName.setBounds(10,160,120,30);
		textfieldDirName.setBounds(120,160,50,30);
		labelDirName2.setBounds(180,160,120,30);
		textfieldDirName.setText(1+"");
		
		this.panel.add(labelDirName);
        this.panel.add(textfieldDirName);
        this.panel.add(labelDirName2);
        
//        JLabel labelColIndex = new JLabel("���ص�");  
//        textfieldColIndex = new JTextField(10);
//		JLabel labelColIndex2 = new JLabel("���ļ��������� , ������");  
//		
//		labelColIndex.setBounds(10,200,120,30);
//		textfieldColIndex.setBounds(120,200,100,30);
//		labelColIndex2.setBounds(230,200,200,30);
//		textfieldColIndex.setText(16+","+17+","+18);
//		
//		this.panel.add(labelColIndex);
//        this.panel.add(textfieldColIndex);
//        this.panel.add(labelColIndex2);
        
        buttonDownload = new JButton("��ʼִ������");
        buttonDownload.setBounds(10,200,570, 45);
        buttonDownload.setEnabled(false);
        buttonDownload.addActionListener(actionListener);
        this.panel.add(buttonDownload);
        
        buttonStop = new JButton("ֹͣ");
        buttonStop.setBounds(10,200,570, 45);
        buttonStop.addActionListener(actionListener);
        this.panel.add(buttonStop);
        
        this.textareaLog = new JTextArea();
        textareaLog.setBounds(10,255,570,300);
        textareaLog.setEditable(false);
        this.panel.add(textareaLog);
	}
	
	ActionListener actionListener = new ActionListener() {
		@Override
		public void actionPerformed(ActionEvent e) {
			if(e.getSource() == buttonFile){
				selectFile = openChoseFile();
				if(selectFile != null){
					textfieldFile.setText(selectFile.getAbsolutePath());
				}
			}else if(e.getSource() == buttonOut){
				selectOutDir = openChoseDir();
				if(selectOutDir != null){
					String path = selectOutDir.getAbsolutePath();
					selectOutDir = new File(path,"phone-data");
					textfieldOut.setText(selectOutDir.getAbsolutePath());
				}
			}else if(e.getSource() == buttonRead){
				if(selectFile == null){
					showMessage("��ѡ���ļ�");
					return;
				}
				if(selectOutDir == null){
					showMessage("��ѡ������Ŀ¼");
					return;
				}
				String nameIndexStr = textfieldDirName.getText();
				try {
					int nameIndex = Integer.parseInt(nameIndexStr);
					//String colIndexStr = textfieldColIndex.getText();
					try {
//						String[] colStrList = colIndexStr.split(",");
//						List<Integer> colList = new ArrayList<Integer>();
//						for (String str : colStrList) {
//							colList.add(Integer.parseInt(str));
//						}
						dataList = readExcelData(selectFile,nameIndex);
						int taskCount = 0;
						downloadCount = 0;
						for (int i=0;i<dataList.size();i++) {
							DownloadData d = dataList.get(i);
							taskCount++;
							for (String src : d.srcList) {
								downloadCount++;
							}
						}
						if(downloadCount > 0){
							buttonDownload.setEnabled(true);
							addLog("��ȡ����"+taskCount+"����,��Ҫ����"+downloadCount+"���ļ�");
						}else{
							showMessage("��ȡ������������Ϊ0");
						}
					} catch (Exception e2) {
						e2.printStackTrace();
						showMessage("���������������ָ�ʽ�������,�ָ�");
					}
				} catch (Exception e2) {
					e2.printStackTrace();
					showMessage("�ļ������������������ָ�ʽ");
				}
				
			}else if(e.getSource() == buttonDownload){
				buttonFile.setEnabled(false);
				buttonRead.setEnabled(false);
				
				buttonDownload.setVisible(false);
				buttonStop.setVisible(true);
				
				new Thread(){
					public void run() {
						startDownload();
					};
				}.start();
			}else if(e.getSource() == buttonStop){
				int result = JOptionPane.showConfirmDialog(null, "ֹͣ���´����ػ��ͷ��ʼ��ȷ��Ҫֹͣ��?", "����",JOptionPane.YES_NO_OPTION);  
				if(result == 0){
					state = 0;
				}
			}
		}
	};
	
	public void startDownload(){
		int tempIndex = 0;
		state = 1;
		for (DownloadData d : dataList) {
			for(int i=0;i<d.srcList.size();i++){
				//ֹͣ
				if(state == 0){
					addLog("ֹͣ����");
					downloadEnd();
					return;
				}
				String src = d.srcList.get(i);
				addLog("("+downloadCount+"/"+(tempIndex+1)+")��ʼ����");
				String fileName = src.substring(src.lastIndexOf('/')+1);  
				String downDirPath = selectOutDir.getAbsolutePath()+File.separator+d.phone;
				boolean success = downLoadFromUrl(src,fileName,downDirPath);
				if(success){
					addLog("���سɹ�SUCCESS��"+d.phone);
				}else{
					addLog("����ʧ��ERROR��"+d.phone);
				}
				tempIndex++;
			}
		}
		showMessage("�������");
		downloadEnd();
	}
	public void downloadEnd(){
		buttonFile.setEnabled(true);
		buttonRead.setEnabled(true);
		buttonDownload.setVisible(true);
		buttonStop.setVisible(false);
	}
	
	public void addLog(String log){
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		String dateStr = sdf.format(new Date());
		logList.addFirst(dateStr+"--"+log);
		if(logList.size() > 500){
			logList.removeLast();
		}
		StringBuffer sb = new StringBuffer();
		for (String str : logList) {
			sb.append(str+"\n");
		}
		this.textareaLog.setText(sb.toString());
	}
	

    public File openChoseFile(){
        JFileChooser fileChooser = new JFileChooser();  
        FileNameExtensionFilter filter = new FileNameExtensionFilter("xls�ļ�", "xls");
        fileChooser.setFileFilter(filter);  
        fileChooser.showDialog(new JLabel(), "ѡ��");  
        File file = fileChooser.getSelectedFile();
        return file;
    }
    
    public File openChoseDir(){
        JFileChooser fileChooser = new JFileChooser();  
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        fileChooser.showDialog(new JLabel(), "ѡ��");  
        File file = fileChooser.getSelectedFile();
        return file;
    }
    
    public void showMessage(String message){
    	JOptionPane.showMessageDialog(null, message, "��ʾ", JOptionPane.INFORMATION_MESSAGE);
    }
    
    public List<DownloadData> readExcelData(File file,Integer nameIndex) throws Exception {
    	List<DownloadData> rowList = new ArrayList<DownloadData>();
    	
    	HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet sheet = wb.getSheetAt(0);
		int lastRowNum = sheet.getLastRowNum();
		
		for(int i=0;i<=lastRowNum;i++){
			HSSFRow rowData = sheet.getRow(i);
			Cell dirNameCell = rowData.getCell(nameIndex-1);
			if(dirNameCell != null){
				
				dirNameCell.setCellType(Cell.CELL_TYPE_STRING);
				String phone = dirNameCell.getStringCellValue();
				
				if(phone != null && !phone.equals("")){
					DownloadData downloadData = new DownloadData();
					downloadData.phone = phone;
					
					List<String> srcList = new ArrayList<String>();
					srcList.add("https://realname.oss-cn-beijing.aliyuncs.com/images/userCardNo/"+phone+"-1.png");
					srcList.add("https://realname.oss-cn-beijing.aliyuncs.com/images/userCardNo/"+phone+"-2.png");
					srcList.add("https://realname.oss-cn-beijing.aliyuncs.com/images/userCardNo/"+phone+"-3.png");
					
					downloadData.srcList = srcList;
					rowList.add(downloadData);
				}
			}
			
		}
    	
    	return rowList;
    }
    
    
    public class DownloadData{
    	public String phone;
    	public List<String> srcList;
    }
    
    public boolean downLoadFromUrl(String urlStr,String fileName,String savePath){  
    	try{
    		URL url = new URL(urlStr);    
            HttpURLConnection conn = (HttpURLConnection)url.openConnection();    
            //���ó�ʱ��Ϊ3��  
            conn.setConnectTimeout(3*1000);  
            //��ֹ���γ���ץȡ������403����  
            conn.setRequestProperty("User-Agent", "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT; DigExt)");  
            //�õ�������  
            InputStream inputStream = conn.getInputStream();    
            //��ȡ�Լ�����  
            byte[] getData = readInputStream(inputStream);     
            //�ļ�����λ��  
            File saveDir = new File(savePath);  
            if(!saveDir.exists()){  
                saveDir.mkdirs();  
            }  
            File file = new File(saveDir+File.separator+fileName);      
            FileOutputStream fos = new FileOutputStream(file);    
            fos.write(getData);   
            if(fos!=null){  
                fos.close();    
            }  
            if(inputStream!=null){  
                inputStream.close();  
            }  
            return true;
    	}catch(Exception e){
    		e.printStackTrace();
    	}
    	return false;  
    }  
    
    public byte[] readInputStream(InputStream inputStream) throws IOException {    
        byte[] buffer = new byte[1024];    
        int len = 0;    
        ByteArrayOutputStream bos = new ByteArrayOutputStream();    
        while((len = inputStream.read(buffer)) != -1) {    
            bos.write(buffer, 0, len);    
        }    
        bos.close();    
        return bos.toByteArray();    
    }   


}
