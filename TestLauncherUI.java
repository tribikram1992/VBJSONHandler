package com.bhs.tests;




import javax.swing.*;
import javax.swing.border.LineBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.*;
import java.awt.*;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.ComponentAdapter;
import java.awt.event.ComponentEvent;
import java.io.File;

public class TestLauncherUI extends JFrame {

    private JTextField environmentField;
    private JTextField browserField;
    private JTextField filePathField;
    private JButton browseButton;

    private JLabel environmentLabel;
    private JLabel browserLabel;
    private JLabel filePathLabel;

    private JTextPane logArea;
    private StyledDocument logDocument;

    private String environment;
    private String browser;
    private String testDataFilePath;

    private JLabel logoLabel;

    public TestLauncherUI() {
        setTitle("Automated Test Launcher");
        setSize(500, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        // Set Nimbus Look and Feel for modern UI
        try {
            UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
        } catch (Exception e) {
            e.printStackTrace();
        }

        setLayout(new BorderLayout(10, 10));

        // Logo Panel with Image
        JPanel logoPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        logoLabel = new JLabel();
        logoPanel.add(logoLabel);
        add(logoPanel, BorderLayout.NORTH);

        // Form Panel (Improved GridBagLayout for better arrangement)
        JPanel formPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);
        gbc.anchor = GridBagConstraints.WEST;

        // Labels and Fields
        environmentLabel = new JLabel("Environment:");
        environmentField = new JTextField(20);
        addFocusListener(environmentField, environmentLabel, "Environment:");

        browserLabel = new JLabel("Browser:");
        browserField = new JTextField(20);
        addFocusListener(browserField, browserLabel, "Browser:");

        filePathLabel = new JLabel("Test Data File Path:");
        filePathField = new JTextField(20);
        filePathField.setEditable(false); // User can't type manually
        browseButton = new JButton("Browse");

        // Adding components to formPanel
        gbc.gridx = 0;
        gbc.gridy = 0;
        formPanel.add(environmentLabel, gbc);
        gbc.gridx = 1;
        formPanel.add(environmentField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 1;
        formPanel.add(browserLabel, gbc);
        gbc.gridx = 1;
        formPanel.add(browserField, gbc);

        gbc.gridx = 0;
        gbc.gridy = 2;
        formPanel.add(filePathLabel, gbc);
        gbc.gridx = 1;
        formPanel.add(filePathField, gbc);
        gbc.gridx = 2;
        formPanel.add(browseButton, gbc);

        // Add formPanel to the center of the frame
        add(formPanel, BorderLayout.CENTER);

        // Bottom Panel (Submit/Clear + Log Area)
        JPanel bottomPanel = new JPanel();
        bottomPanel.setLayout(new BoxLayout(bottomPanel, BoxLayout.Y_AXIS));

        JPanel buttonPanel = new JPanel();
        JButton submitButton = new JButton("Submit");
        JButton clearButton = new JButton("Clear Messages");
        buttonPanel.add(submitButton);
        buttonPanel.add(clearButton);
        bottomPanel.add(buttonPanel);

        logArea = new JTextPane();
        logArea.setEditable(false);
        logDocument = logArea.getStyledDocument();
        logArea.setBorder(BorderFactory.createTitledBorder("Messages"));

        JScrollPane logScroll = new JScrollPane(logArea) {
            @Override
            public Dimension getPreferredSize() {
                int lines = logArea.getDocument().getDefaultRootElement().getElementCount();
                int height = Math.min(lines * 20 + 40, 250);
                return new Dimension(super.getPreferredSize().width, height);
            }
        };
        bottomPanel.add(logScroll);

        add(bottomPanel, BorderLayout.SOUTH);

        // Browse Button Action (with Excel file filter)
        browseButton.addActionListener(e -> {
            JFileChooser fileChooser = new JFileChooser();

            // Set file filter for Excel-readable formats
            FileNameExtensionFilter filter = new FileNameExtensionFilter(
                    "Excel Files (*.xls, *.xlsx, *.csv)", "xls", "xlsx", "csv");
            fileChooser.setFileFilter(filter);

            // Show the system native file chooser
            int result = fileChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                String fileName = selectedFile.getName().toLowerCase();

                if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx") || fileName.endsWith(".csv")) {
                    filePathField.setText(selectedFile.getAbsolutePath());
                    resetValidation(filePathField, filePathLabel, "Test Data File Path:");
                } else {
                    appendMessage("Selected file is not a valid Excel file (.xls, .xlsx, .csv).\n", Color.RED);
                }
            }
        });

        // Submit Action
        submitButton.addActionListener(e -> {
            boolean isValid = true;
            clearLog();

            environment = environmentField.getText().trim();
            browser = browserField.getText().trim();
            testDataFilePath = filePathField.getText().trim();

            if (environment.isEmpty()) {
                markInvalid(environmentField, environmentLabel, "Environment:");
                isValid = false;
            }
            if (browser.isEmpty()) {
                markInvalid(browserField, browserLabel, "Browser:");
                isValid = false;
            }
            if (testDataFilePath.isEmpty()) {
                markInvalid(filePathField, filePathLabel, "Test Data File Path:");
                isValid = false;
            }

            if (isValid) {
                File file = new File(testDataFilePath);
                if (!file.exists()) {
                    appendMessage("Provided file does not exist. Please provide a valid test data file.\n", Color.RED);
                } else {
                    appendMessage("Inputs submitted successfully:\n", new Color(0, 128, 0));
                    appendMessage("Environment: " + environment + "\n", Color.BLACK);
                    appendMessage("Browser: " + browser + "\n", Color.BLACK);
                    appendMessage("File Path: " + testDataFilePath + "\n", Color.BLACK);
                }
            } else {
                appendMessage("Please fill all required fields.\n", Color.RED);
            }

            logArea.revalidate();
            logArea.repaint();
        });

        // Clear Button Action
        clearButton.addActionListener(e -> clearLog());

        // Add window resize listener to adjust the logo size dynamically
        addComponentListener(new ComponentAdapter() {
            public void componentResized(ComponentEvent e) {
                updateLogoSize();
            }
        });

        setVisible(true);
    }

    // Update logo size dynamically based on window size
    private void updateLogoSize() {
        // Assuming the logo is located at a relative path (update the path as needed)
        ImageIcon icon = new ImageIcon("C:\\baxterCode\\Baxter.png");

        // Calculate dynamic size based on window dimensions
        int width = getWidth();
        int height = getHeight();

        // Set logo size to 10% of the width and 5% of the height
        int logoWidth = (int) (width * 0.1);
        int logoHeight = (int) (height * 0.05);

        // Resize the logo
        Image resizedImage = icon.getImage().getScaledInstance(logoWidth, logoHeight, Image.SCALE_SMOOTH);
        icon = new ImageIcon(resizedImage);

        // Set the logo on the label
        logoLabel.setIcon(icon);
    }

    private void markInvalid(JTextField field, JLabel label, String labelText) {
        label.setText(labelText + " ✱");
        label.setForeground(Color.RED);
        field.setBorder(new LineBorder(Color.RED, 2));
    }

    private void resetValidation(JTextField field, JLabel label, String labelText) {
        label.setText(labelText);
        label.setForeground(Color.BLACK);
        field.setBorder(UIManager.getLookAndFeel().getDefaults().getBorder("TextField.border"));
    }

    private void addFocusListener(JTextField field, JLabel label, String labelText) {
        field.addFocusListener(new FocusAdapter() {
            @Override
            public void focusLost(FocusEvent e) {
                if (!field.getText().trim().isEmpty()) {
                    resetValidation(field, label, labelText);
                }
            }
        });
    }

    private void appendMessage(String message, Color color) {
        StyleContext context = new StyleContext();
        AttributeSet attrSet = context.addAttribute(SimpleAttributeSet.EMPTY, StyleConstants.Foreground, color);
        try {
            logDocument.insertString(logDocument.getLength(), message, attrSet);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void clearLog() {
        logArea.setText("");
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(TestLauncherUI::new);
    }
}
