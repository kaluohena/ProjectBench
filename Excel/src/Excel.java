import java.awt.BorderLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;



public class Excel extends JPanel implements ActionListener {
	
	static private final String newline = "\n";
    JButton openButton, saveButton, calculateCloseButton, calculateOpenButton, timeConnectionButton;
    JTextArea log;
    JFileChooser fc;
    String filePath;
    String outString;
    
    public Excel() {
    	super(new BorderLayout());

        //Create the log first, because the action listeners
        //need to refer to it.
        log = new JTextArea(5,20);
        log.setMargin(new Insets(5,5,5,5));
        log.setEditable(false);
        JScrollPane logScrollPane = new JScrollPane(log);

        //Create a file chooser
        fc = new JFileChooser();

        //Uncomment one of the following lines to try a different
        //file selection mode.  The first allows just directories
        //to be selected (and, at least in the Java look and feel,
        //shown).  The second allows both files and directories
        //to be selected.  If you leave these lines commented out,
        //then the default mode (FILES_ONLY) will be used.
        //
        //fc.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //fc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES);

        //Create the open button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        openButton = new JButton("Open a File...",
                                 createImageIcon("assets/Open16.gif"));
        openButton.addActionListener(this);

        //Create the save button.  We use the image from the JLF
        //Graphics Repository (but we extracted it from the jar).
        /*saveButton = new JButton("Save a File...",
                                 createImageIcon("images/Save16.gif"));
        saveButton.addActionListener(this);*/
        
        calculateCloseButton = new JButton("Calculate Close");
        calculateCloseButton.addActionListener(this);
        
        calculateOpenButton = new JButton("Calculate Open");
        calculateOpenButton.addActionListener(this);
        
        //Create the time connection button 
        timeConnectionButton = new JButton("Copy and Paste");
        timeConnectionButton.addActionListener(this);

        //For layout purposes, put the buttons in a separate panel
        JPanel buttonPanel = new JPanel(); //use FlowLayout
        buttonPanel.add(openButton);
        //buttonPanel.add(saveButton);
        buttonPanel.add(calculateCloseButton);
        buttonPanel.add(calculateOpenButton);
        buttonPanel.add(timeConnectionButton);

        //Add the buttons and the log to this panel.
        add(buttonPanel, BorderLayout.PAGE_START);
        add(logScrollPane, BorderLayout.CENTER);
    }
    
    public void actionPerformed(ActionEvent e) {
    	// TODO Auto-generated method stub
    	//Handle open button action.
    	if (e.getSource() == openButton) {
    		if (outString != null) {
    			fc.setCurrentDirectory(new File(outString));
    		} else {
    			File startFile = new File(System.getProperty("user.home") + "/Desktop");
    			fc.setCurrentDirectory(startFile);
    		}
    		int returnVal = fc.showOpenDialog(Excel.this);

    		if (returnVal == JFileChooser.APPROVE_OPTION) {
    			File file = fc.getSelectedFile();
    			filePath = file.getPath();
    			// store the file location for staying the current position when open the file system next time
    			outString = fc.getSelectedFile().getPath();
    			//This is where a real application would open the file.
    			log.append("Opening: " + file.getPath()+ "." + newline);
    		} else {
    			log.append("Open command cancelled by user." + newline);
    		}
    		log.setCaretPosition(log.getDocument().getLength());

    		//Handle save button action.
    	} /*else if (e.getSource() == saveButton) {
            int returnVal = fc.showSaveDialog(FileChooserDemo.this);
            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File file = fc.getSelectedFile();
                //This is where a real application would save the file.
                log.append("Saving: " + file.getName() + "." + newline);
            } else {
                log.append("Save command cancelled by user." + newline);
            }
            log.setCaretPosition(log.getDocument().getLength());

        }*/ else if (e.getSource() == calculateCloseButton) {
        	if (filePath != null) {
        		ReadWrite readWriteExcel = new ReadWrite(filePath, "c:/temp/CaculateCloseResult.xls", log, false);
        		try {
        			readWriteExcel.readWrite();
        		}catch (Throwable t) {
        			t.printStackTrace();
        		}

        	} 
        } else if(e.getSource() == calculateOpenButton) {
        	if (filePath != null) {
        		ReadWrite readWriteExcel = new ReadWrite(filePath, "c:/temp/CaculateOpenResult.xls", log, true);
        		try {
        			readWriteExcel.readWrite();
        		}catch (Throwable t) {
        			t.printStackTrace();
        		}
        	}
        } else if (e.getSource() == timeConnectionButton) {        		
        	if (filePath != null) {
        		CopyPaste copyPasteExcel = new CopyPaste(filePath, "c:/temp/PastedExample.xls", log);
        		try {
        			copyPasteExcel.copyPaste();
        		}catch (Throwable t) {
        			t.printStackTrace();
        		}
        	}
        } else {
        	log.append("no file selected");
        }
    }
	
    
    /** Returns an ImageIcon, or null if the path was invalid. */
    protected static ImageIcon createImageIcon(String path) {
        java.net.URL imgURL = Excel.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }
    
    /**
     * Create the GUI and show it.  For thread safety,
     * this method should be invoked from the
     * event dispatch thread.
     */
    private static void createAndShowGUI() {
        //Create and set up the window.
        JFrame frame = new JFrame("ExcelProcessingToolV4.3");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        //Add content to the window.
        frame.add(new Excel());

        //Display the window.
        frame.pack();
        frame.setVisible(true);
    }
    
    public static void main(String[] args) {
        //Schedule a job for the event dispatch thread:
        //creating and showing this application's GUI.
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                //Turn off metal's use of bold fonts
                UIManager.put("swing.boldMetal", Boolean.FALSE); 
                createAndShowGUI();
            }
        });
    }


}
