package imagepull;

import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.commons.io.FileUtils;

import net.lingala.zip4j.core.ZipFile;
import net.lingala.zip4j.exception.ZipException;

@SuppressWarnings("serial")
public class Pull extends JFrame implements ActionListener{

	public static void main(String[] args) {
		Pull view = new Pull();
		view.setVisible(true);
	}
	
	File doc;
	JFileChooser fileChooser;
	JLabel label;

	public Pull(){
		
		setTitle("Image Pull");
		setSize(280, 80);
		setMinimumSize(getSize());
		
		setDefaultCloseOperation(3);
		setLocationRelativeTo(null);//center of screen
		
		setLayout(new BorderLayout());
		
		JButton button = new JButton();
		button.setText("Select File");
		button.setSize(100, 50);
		button.setFocusable(false);
		button.addActionListener(this);
		
		label = new JLabel();
		label.setText("");

		JPanel top = new JPanel();
		top.add(button);
		
		JPanel bottom = new JPanel();
		bottom.add(label);
		
		add(top, BorderLayout.NORTH);
		add(bottom, BorderLayout.SOUTH);
	}
	
	@Override
	public void actionPerformed(ActionEvent e) {
		fileChooser = new JFileChooser();
		fileChooser.setCurrentDirectory(null);
		
		FileNameExtensionFilter filter = new FileNameExtensionFilter("Office Files", "docx", "pptx", "potx");
		fileChooser.setFileFilter(filter);
		
		int action = fileChooser.showOpenDialog(this);
		
		if(action != JFileChooser.APPROVE_OPTION){
			label.setText("");
			return;
		}
		
		doc = fileChooser.getSelectedFile();
		
		System.out.println(doc.getName());
		
		File selectedFile = fileChooser.getSelectedFile();
		
		try{
			File zipCopy = new File(selectedFile.getParent() + "/" + selectedFile.getName() + ".zip");
			FileUtils.copyFile(selectedFile, zipCopy);
			
			String folderPath = selectedFile.getParent() + "/" + selectedFile.getName().split("\\.")[0];
			File newFolder = new File(folderPath);
			
			ZipFile z = new ZipFile(zipCopy);
			z.extractAll(folderPath);
			
			// File cleanup
			zipCopy.delete();
			
			File media = new File(folderPath + "/word/media");
			
			if(!media.exists()){
				System.out.println("'word' does not exist. Changing to 'ppt'");
				media = new File(folderPath + "/ppt/media");
			}
			
			System.out.println("Name to be saved as:    "+ doc.getName().split("\\.")[0]);
			//File newMedia = new File(selectedFile.getParent() + "/" + doc.getName().split("\\.")[0]);
			File newMedia = new File(selectedFile.getParent() + "/" + doc.getName() + " - Media");
			
			
			// Copies entire media folder to the new media folder
			FileUtils.copyDirectory(media, newMedia);
			
			// File cleanup
			FileUtils.deleteDirectory(newFolder);
			
			label.setText("Images extracted to current directory.");
		} catch(IOException | ZipException o){
			o.printStackTrace();
		}
	}
}
