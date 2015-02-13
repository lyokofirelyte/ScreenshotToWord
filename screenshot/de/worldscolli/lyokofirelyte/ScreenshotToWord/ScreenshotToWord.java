package de.worldscolli.lyokofirelyte.ScreenshotToWord;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.InputStream;

import javax.imageio.ImageIO;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.UIManager;

import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.wml.Drawing;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;

import lombok.SneakyThrows;

public class ScreenshotToWord extends JFrame {

	public ScreenshotToWord(){
		start();
	}
	
	@SneakyThrows
	public static void main(String[] args){
		UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		new ScreenshotToWord();
	}
	
	JFrame frame;
	JTextField field;
	
	private int startX;
	private int startY;
	
	private int width;
	private int height;
	
	@SneakyThrows
	public void add(File file){
		
		File output = new File(field.getText());
        WordprocessingMLPackage wordprocessingMLPackage = WordprocessingMLPackage.load(output);

        InputStream inputStream = new java.io.FileInputStream(file);
        long fileLength = file.length();    

        byte[] bytes = new byte[(int)fileLength];

        int offset = 0;
        int numRead = 0;

        while (offset < bytes.length
               && (numRead = inputStream.read(bytes, offset, bytes.length-offset)) >= 0) {
            offset += numRead;
        }

        inputStream.close();
        
        P p = newImage(wordprocessingMLPackage, bytes, null, null, 0, 1);

        wordprocessingMLPackage.getMainDocumentPart().addObject(p);
        wordprocessingMLPackage.save(output);
    }

    public P newImage( WordprocessingMLPackage wordMLPackage, byte[] bytes, 
        String filenameHint, String altText, int id1, int id2) throws Exception {
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(wordMLPackage, bytes);
        Inline inline = imagePart.createImageInline( filenameHint, altText, id1, id2);

        ObjectFactory factory = new ObjectFactory();

        P  p = factory.createP();
        R  run = factory.createR();

        p.getParagraphContent().add(run);        
        Drawing drawing = factory.createDrawing();      
        run.getRunContent().add(drawing);       
        drawing.getAnchorOrInline().add(inline);

        return p;
    }  
	
	private ActionListener listener = new ActionListener(){
		
		@Override @SneakyThrows
		public void actionPerformed(ActionEvent e) {
			
			switch (e.getActionCommand()){
			
				case "capture":
					
					setTitle("Saving...");
					
					try {
						Rectangle screenRect = new Rectangle(startX, startY, width, height);
						BufferedImage capture = new Robot().createScreenCapture(screenRect);
						ImageIO.write(capture, "png", new File("temp.png"));
						
						if (new File("temp.png").exists()){
							add(new File("temp.png"));
						}
					} catch (Exception ee){}
					
					setTitle("Ready!");
					
				break;
				
				case "select":
					
					frame = new JFrame();
					frame.setTitle("Selection Window");
					
					JPanel panel = new JPanel();
					panel.setPreferredSize(new Dimension(400, 300));
					panel.setLayout(new FlowLayout(FlowLayout.CENTER, 0, 0));
					panel.setOpaque(false);
					panel.setBackground(new Color(0, 0, 0, 0.3f));
					
					JLabel label = new JLabel("Size this window to represent your capture area.\n Press save when complete.");
					label.setOpaque(false);
					label.setBackground(new Color(0, 0, 0, 0.3f));
					panel.add(label);
					
					JButton button = new JButton("Save!");
					button.addActionListener(listener);
					button.setActionCommand("confirmSelection");
					panel.add(button);
					
					frame.add(panel);
					
					frame.setLocationRelativeTo(null);
					frame.setAlwaysOnTop(true);
					frame.pack();
					frame.setVisible(true);
					frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
					
				break;
				
				case "confirmSelection":
					
					startX = new Integer(frame.getX());
					startY = new Integer(frame.getY());
					width = new Integer(frame.getWidth());
					height = new Integer(frame.getHeight());
					frame.dispose();
					
				break;
			}
		}
	};
	
	public void start(){
		
		setTitle("ScreenshotToWord (ITS 260)");
		JPanel mainPanel = new JPanel();
		mainPanel.setLayout(new GridLayout(0, 1));
		
		JPanel textPanel = new JPanel();
		textPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 0, 0));
		textPanel.setBorder(BorderFactory.createTitledBorder("<html><div style='color: 046344'>" + "Path Selection" + "</div></html>"));
		
		JPanel buttonHolderPanel = new JPanel();

		JPanel buttonPanel = new JPanel();
		buttonPanel.setLayout(new FlowLayout(FlowLayout.CENTER, 0, 0));
		
		JPanel buttonPanel2 = new JPanel();
		buttonPanel2.setLayout(new FlowLayout(FlowLayout.CENTER, 0, 0));
		
		field = new JTextField("Name of word doc");
		field.setPreferredSize(new Dimension(200, 20));
		textPanel.add(field);
		
		JButton button = new JButton("Capture!");
		button.addActionListener(listener);
		button.setActionCommand("capture");
		buttonPanel.setBorder(BorderFactory.createTitledBorder("<html><div style='color: 046344'>" + "" + "</div></html>"));
		buttonPanel.add(button);
		
		JButton button2 = new JButton("Select Window...");
		button2.addActionListener(listener);
		button2.setActionCommand("select");
		buttonPanel2.setBorder(BorderFactory.createTitledBorder("<html><div style='color: 046344'>" + "" + "</div></html>"));
		buttonPanel2.add(button2);
		
		buttonHolderPanel.add(buttonPanel);
		buttonHolderPanel.add(buttonPanel2);
		mainPanel.add(textPanel);
		mainPanel.add(buttonHolderPanel);
		
		add(mainPanel);
		
		setLocationRelativeTo(null);
		setAlwaysOnTop(true);
		pack();
		setVisible(true);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setResizable(false);
	}
}