import java.awt.Desktop;
import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.IOException;
import java.net.URISyntaxException;

import jxl.read.biff.BiffException;

import javax.swing.JLabel;
import javax.swing.JTextField;
import java.awt.GridLayout;
import javax.swing.JTextArea;

public class ExcelMain extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	private String[][] strArrTest;
	private JLabel lblFileSon, lblFileDad;
	private JButton btnColor;

	private String strFileSon;
	private String strFileDad;
	private JTextField textFieldDadRow;
	private JTextField textFieldDupliRow;
	JTextField textFieldSonCol;
	JTextField textFieldSonRow;
	JTextField textFieldDadCol;
	JTextField textFieldDupliCol;
	
	String strFileDuplicate;
	JLabel lblDupliFileName;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ExcelMain frame = new ExcelMain();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public void showDadDlg() {
		JFileChooser fileChoose = new JFileChooser();
		int nChooser = fileChoose.showOpenDialog(null);
		if (nChooser == JFileChooser.APPROVE_OPTION) {
			strFileDad = fileChoose.getSelectedFile().getPath();
			lblFileDad.setText(fileChoose.getSelectedFile().getName());
		} else {
			strFileDad = null;
			return;
		}
//		
//		FileDialog fileDadDlg = new FileDialog(this, "打开父文件", FileDialog.LOAD);
//		fileDadDlg.setVisible(true);
//		String strDad = fileDadDlg.getFile();
//		if (strDad == null)
//			strFileDad = null;
//		else {
//			lblFileDad.setText(strDad);
//			strFileDad = fileDadDlg.getDirectory() + strDad;
//		}
	}

	public void showSunDlg() {
//		fileSunDlg = new FileDialog(this, "打开子文件", FileDialog.LOAD);
//		fileSunDlg.setVisible(true);
//		String strSun = fileSunDlg.getFile();
//		if (strSun == null)
//			strFileSon = null;
//		else {
//			lblFileSon.setText(strSun + "-->");
//			strFileSon = fileSunDlg.getDirectory() + strSun;
//		}

		
		JFileChooser fileChoose = new JFileChooser();
		int nChooser = fileChoose.showOpenDialog(null);
		if (nChooser == JFileChooser.APPROVE_OPTION) {
			lblFileSon.setText(fileChoose.getSelectedFile().getName());
			strFileSon = fileChoose.getSelectedFile().getPath();
		} else {
			strFileSon = null;
			return;
		}
	}

	/**
	 * Create the frame.
	 */
	public ExcelMain() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 573, 457);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(new GridLayout(0, 1, 0, 0));

		JPanel paneltop = new JPanel();
		contentPane.add(paneltop);

		JPanel panel_Son = new JPanel();
		paneltop.add(panel_Son);
		paneltop.setBorder(new TitledBorder("查看两个表中是否存在着重复数据"));
		JButton btnSun = new JButton("打开子文件");
		panel_Son.add(btnSun);

		JLabel label_2 = new JLabel("开始列");
		panel_Son.add(label_2);

		textFieldSonCol = new JTextField();
		textFieldSonCol.setDocument(new NumberLenghtLimitedDmt());
		panel_Son.add(textFieldSonCol);
		textFieldSonCol.setColumns(10);

		JLabel label_3 = new JLabel("开始行");
		panel_Son.add(label_3);

		textFieldSonRow = new JTextField();
		textFieldSonRow.setDocument(new NumberLenghtLimitedDmt());
		panel_Son.add(textFieldSonRow);
		textFieldSonRow.setColumns(10);

		lblFileSon = new JLabel("子文件名");
		panel_Son.add(lblFileSon);
		btnSun.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				showSunDlg();
				if (strFileSon != null) {
					if (strFileDad != null)
						btnColor.setEnabled(true);
				}
			}
		});

		JPanel panelDad = new JPanel();
		paneltop.add(panelDad);

		JButton btnFather = new JButton("打开父文件");
		btnFather.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				showDadDlg();
				if (strFileDad != null) {
					if (strFileSon != null)
						btnColor.setEnabled(true);
				}
			}
		});
		panelDad.add(btnFather);

		JLabel label_4 = new JLabel("开始列");
		panelDad.add(label_4);

		textFieldDadCol = new JTextField();
		textFieldDadCol.setDocument(new NumberLenghtLimitedDmt());
		textFieldDadCol.setColumns(10);
		panelDad.add(textFieldDadCol);

		JLabel label_5 = new JLabel("开始行");
		panelDad.add(label_5);

		textFieldDadRow = new JTextField();
		textFieldDadRow.setDocument(new NumberLenghtLimitedDmt());
		textFieldDadRow.setColumns(10);
		panelDad.add(textFieldDadRow);

		lblFileDad = new JLabel("父文件名");
		panelDad.add(lblFileDad);

		btnColor = new JButton("color");
		paneltop.add(btnColor);
		btnColor.setEnabled(false);

		btnColor.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				ExcelUtils excel = new ExcelUtils();
				try {
					String strSonRow=textFieldSonRow.getText().trim();
					String strSonCol=textFieldSonCol.getText().trim();
					String strDadRow = textFieldDadRow.getText().trim();
					String strDadCol = textFieldDadCol.getText().trim();
					if (strSonCol.isEmpty() || strSonRow.isEmpty()
							|| strDadCol.isEmpty() || strDadRow.isEmpty())
					{
						JOptionPane.showMessageDialog(null, "请先打开数据文件，并将其他数据填写完整");
						return;
					}
					int nSonRow=Integer.parseInt(strSonRow);
					int nSonCol=Integer.parseInt(strSonCol);
					int nDadRow=Integer.parseInt(strDadRow);
					int nDadCol=Integer.parseInt(strDadCol);
					if(nSonRow==0||nSonCol==0||nDadRow==0||nDadCol==0)
						return;
					nSonRow--;
					nSonCol--;
					nDadRow--;
					nDadCol--;
					if (excel.ReadString(strFileSon, nSonRow, nSonCol)) {
						textFieldSonRow.setText(" ");
						textFieldSonCol.setText(" ");
						textFieldDadRow.setText(" ");
						textFieldDadCol.setText(" ");
						strArrTest = excel.getStringData().get(0);
						excel.colorIt(strFileDad, strArrTest, nDadRow, 
								nDadCol, nSonRow,nSonCol,1);
						Desktop desktop = Desktop.getDesktop();
						File fileOpen = new File(strFileDad);
						desktop.open(fileOpen);
					}
				} catch (BiffException e1) {
					e1.printStackTrace();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});

		JTextArea textArea = new JTextArea();
		textArea.setText("如果父文件中指定列包含子文件指定列的内容，将父文件中相应的项设置为红色。");
		paneltop.add(textArea);

		JPanel panelBottom = new JPanel();
		contentPane.add(panelBottom);

		panelBottom.setBorder(new TitledBorder("查看表中是否有重复数据"));
		JButton btnOpenDupliFile = new JButton("打开源文件");
		panelBottom.add(btnOpenDupliFile);
		btnOpenDupliFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser fileChoose = new JFileChooser();
				int nChooser = fileChoose.showOpenDialog(null);
				if (nChooser == JFileChooser.APPROVE_OPTION) {
					lblDupliFileName.setText(fileChoose.getSelectedFile().getName());
					strFileDuplicate = fileChoose.getSelectedFile().getPath();
				} else {
					strFileDuplicate = null;
					return;
				}
			}
		});
		JLabel label = new JLabel("要检查的列");
		panelBottom.add(label);

		textFieldDupliCol = new JTextField();
		textFieldDupliCol.setDocument(new NumberLenghtLimitedDmt());
		panelBottom.add(textFieldDupliCol);
		textFieldDupliCol.setColumns(10);

		JLabel label_1 = new JLabel("起始行");
		panelBottom.add(label_1);

		textFieldDupliRow = new JTextField();
		textFieldDupliRow.setDocument(new NumberLenghtLimitedDmt());
		panelBottom.add(textFieldDupliRow);
		textFieldDupliRow.setColumns(10);
		
		lblDupliFileName = new JLabel("文件名");
		panelBottom.add(lblDupliFileName);

		JButton btnDuplicate = new JButton("开始检查");
		panelBottom.add(btnDuplicate);
		btnDuplicate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				ExcelUtils excel = new ExcelUtils();
				try {
					String strDupCol = textFieldDupliCol.getText();
					String strDupRow=textFieldDupliRow.getText();
					if (strDupCol.isEmpty() || strDupCol == null
							||strFileDuplicate.isEmpty()||strFileDuplicate==null
							||strDupRow.isEmpty()||strDupRow==null)
					{
						JOptionPane.showMessageDialog(null, "请先打开数据文件，并将其他数据填写完整");
						return;
					}
					int nCol = Integer.parseInt(strDupCol)-1;
					int nRow=Integer.parseInt(strDupRow)-1;
					excel.isDuplicate(strFileDuplicate, nCol, nRow);
					Desktop desktop = Desktop.getDesktop();
					File fileOpen = new File(strFileDuplicate);
					desktop.open(fileOpen);
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			}
		});

		JPanel panelMid = new JPanel();
		contentPane.add(panelMid);

		JButton btnResult = new JButton("检查摇号结果");
		btnResult.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String strResult = getCurrentDir() + "result.xls";
				String strTest = getCurrentDir() + "test.xls";
				ExcelUtils myexcel = new ExcelUtils();
				myexcel.CheckResult(strTest, strResult);
				Desktop desktop = Desktop.getDesktop();
				File fileOpen = new File(strResult);
				try {
					desktop.open(fileOpen);
				} catch (IOException e) {
					e.printStackTrace();
				}

			}
		});
		panelMid.add(btnResult);
	}

	public String getCurrentDir() {
		String curdir = null;
		try {
			curdir = Thread.currentThread().getContextClassLoader()
					.getResource("").toURI().getPath();
		} catch (URISyntaxException e) {
			e.printStackTrace();
		}
		return curdir;

	}
}

class Filter extends javax.swing.filechooser.FileFilter {
	public boolean accept(File f) {
		if (f.isDirectory()) {
			return true;
		}
		return f.getName().endsWith(".java");
	}

	public String getDescription() {
		return "*.java";
	}

}