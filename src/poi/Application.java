package poi;

import java.awt.AWTException;
import java.awt.Button;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.EmptyBorder;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;

@SuppressWarnings("serial")
public class Application extends JFrame {
	
	static JLabel lblNewLabel ;
	static JProgressBar progressBar;
	static JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Application frame = new Application();
					frame.setVisible(true);
					frame.setTitle("Chargebacks Tool");
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Application() {
		
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		int hieght = screenSize.height/2;
		int width = (screenSize.width/2) - 100;
		
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, width, hieght);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);

		JPanel panel = new JPanel();
		panel.setBounds(5, 5, 639, 380);
		
			contentPane.add(panel);
			panel.setLayout(null);
				
			JButton btnStart = new JButton("COMPILE PDF");
			btnStart.setBackground(UIManager.getColor("Button.background"));
			btnStart.setBounds(229, 38, 180, 33);
			btnStart.setFont(new Font("Calibri", Font.PLAIN, 20));
			btnStart.setForeground(Color.BLACK);
			panel.add(btnStart);
			
	    progressBar = new JProgressBar();
	    progressBar.setBounds(177, 100, 285, 11);
	    progressBar.setVisible(false);
	    panel.add(progressBar);
	    
	    lblNewLabel = new JLabel("Progress: ");
	    lblNewLabel.setBounds(256, 121, 126, 23);
	    lblNewLabel.setVisible(false);
	    panel.add(lblNewLabel);
	    
	    JButton btnFioriTemplate = new JButton("FIORI TEMPLATE");
	    btnFioriTemplate.setBounds(424, 228, 166, 25);
	    btnFioriTemplate.addActionListener(new ActionListener() {
	    	public void actionPerformed(ActionEvent e) {
	    		try {
	    		Fiori.main();
	    		}catch (BiffException| IOException| WriteException| UnsupportedLookAndFeelException| 
	    				InvalidFormatException| ClassNotFoundException| InstantiationException| IllegalAccessException e2) {}
	    	}
	    });
	    btnFioriTemplate.setFont(new Font("Arial", Font.PLAIN, 14));
	    panel.add(btnFioriTemplate);
	    
	    JButton btnMarkItemsClosed = new JButton("MARK ITEMS CLOSED");
	    btnMarkItemsClosed.setBounds(36, 230, 180, 23);
	    btnMarkItemsClosed.addActionListener(new ActionListener() {
	    	public void actionPerformed(ActionEvent e) {
	    		try {
	    			poi.ChangetoClosed.main();
	    		} catch (InvalidFormatException | BiffException | WriteException | ClassNotFoundException
	    				| InstantiationException | IllegalAccessException | IOException
	    				| UnsupportedLookAndFeelException e1) {
	    			// TODO Auto-generated catch block
	    			e1.printStackTrace();
	    		} catch (InterruptedException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	    	}
	    });
	    btnMarkItemsClosed.setFont(new Font("Arial", Font.PLAIN, 12));
	    panel.add(btnMarkItemsClosed);
	    
	    Button button = new Button("QUIT");
	    button.setFont(new Font("Calibri", Font.PLAIN, 10));
	    button.addActionListener(new ActionListener() {
	    	public void actionPerformed(ActionEvent e) {
	    		System.exit(0);
	    	}
	    });
	    button.setBounds(292, 328, 54, 21);
	    panel.add(button);
		
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				progressBar.setVisible(true);
				lblNewLabel.setVisible(true);
				
					try {
						poi.ChargeBacks.main();
					} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
							| NullPointerException | InvalidFormatException | IOException | AWTException
							| UnsupportedLookAndFeelException | InterruptedException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					} catch (Exception e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
			}
		});
	}
}
