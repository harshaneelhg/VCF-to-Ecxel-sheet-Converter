package source;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class VCFToExcelSheetConverter extends javax.swing.JFrame {
    
    private String dir;
    private String src;

    public VCFToExcelSheetConverter() {
        initComponents();
        txtSrcFile.setEditable(false);
        txtDestFile.setEditable(false);
    }

    @SuppressWarnings("unchecked")
    private void initComponents() {

        txtSrcFile = new javax.swing.JTextField();
        btnBrowseFile = new javax.swing.JButton();
        lblSrcFile = new javax.swing.JLabel();
        lblDestFile = new javax.swing.JLabel();
        txtDestFile = new javax.swing.JTextField();
        btnBrowseDestFile = new javax.swing.JButton();
        btnStart = new javax.swing.JButton();
        lblFileName = new javax.swing.JLabel();
        txtFileName = new javax.swing.JTextField();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        btnBrowseFile.setText("Browse");
        btnBrowseFile.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnBrowseFileActionPerformed(evt);
            }
        });

        lblSrcFile.setText("Choose your contact file : ");

        lblDestFile.setText("Choose Destination : ");

        btnBrowseDestFile.setText("Browse");
        btnBrowseDestFile.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnBrowseDestFileActionPerformed(evt);
            }
        });

        btnStart.setText("Finish");
        btnStart.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnStartActionPerformed(evt);
            }
        });

        lblFileName.setText("Filename :");

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addGap(19, 19, 19)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(lblDestFile)
                            .addComponent(lblSrcFile)
                            .addComponent(lblFileName))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(txtSrcFile, javax.swing.GroupLayout.PREFERRED_SIZE, 232, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(18, 18, 18)
                                .addComponent(btnBrowseFile))
                            .addGroup(layout.createSequentialGroup()
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                                    .addComponent(txtFileName, javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(txtDestFile, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 232, Short.MAX_VALUE))
                                .addGap(18, 18, 18)
                                .addComponent(btnBrowseDestFile))))
                    .addGroup(layout.createSequentialGroup()
                        .addGap(217, 217, 217)
                        .addComponent(btnStart)))
                .addContainerGap(34, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(108, 108, 108)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(txtSrcFile, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnBrowseFile)
                    .addComponent(lblSrcFile))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblDestFile)
                    .addComponent(txtDestFile, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(btnBrowseDestFile))
                .addGap(26, 26, 26)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(lblFileName)
                    .addComponent(txtFileName, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 33, Short.MAX_VALUE)
                .addComponent(btnStart)
                .addGap(22, 22, 22))
        );

        pack();
    }

    private void btnBrowseFileActionPerformed(java.awt.event.ActionEvent evt) {

        JFrame jf = new JFrame();
        jf.setAlwaysOnTop(true);
        String filename = "";
        JFileChooser chooser = new JFileChooser();
        chooser.setFileFilter(new VCFFileFilter());
        chooser.setOpaque(true);
        chooser.setLocation(200, 200);
        jf.setVisible(true);
        jf.setVisible(false);
        int val = chooser.showOpenDialog(jf);
        if (val == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();

            try {
                filename = file.getAbsolutePath();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Exception occured : " + e + "/nCheck File paths.");
            }
        }
        src = filename;
        System.out.println(filename);
        txtSrcFile.setText(filename);
    }

    private void btnBrowseDestFileActionPerformed(java.awt.event.ActionEvent evt) {
        JFrame jf = new JFrame();
        jf.setAlwaysOnTop(true);
        String dirname = "";
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setOpaque(true);
        chooser.setLocation(200, 200);
        jf.setVisible(true);
        jf.setVisible(false);
        int val = chooser.showOpenDialog(jf);
        if (val == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();

            try {
                dirname = file.getAbsolutePath();
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "Exception occured : " + e + "/nCheck File paths.");
            }
        }
        dir = dirname;
        System.out.println(dirname);
        txtDestFile.setText(dirname);
    }

    private void btnStartActionPerformed(java.awt.event.ActionEvent evt) {

        String f = txtFileName.getText();
        try {
            File exlFile = new File(dir + "/" + f + ".xls");
            WritableWorkbook writableWorkbook = Workbook.createWorkbook(exlFile);
            WritableSheet writableSheet = writableWorkbook.createSheet("Sheet1", 0);
            
            File file = new File(src);
            BufferedReader br = new BufferedReader(new FileReader(file));
            String s = br.readLine();
            int i = 0;
            while (s != null) {
                int j = 0;
                if (s.contains("FN:")) {
                    System.out.println(s);
                    Label l = new Label(j, i, s.substring(3, s.length()));
                    j++;

                    String s1 = br.readLine();
                    boolean flag = true;
                    while (s1.contains("TEL")) {
                        if (flag) {
                            writableSheet.addCell(l);
                            flag = false;
                        }
                        if (s1.contains("PREF")) {
                            Label l1 = new Label(j, i, s1.substring(14, s1.length()));
                            j++;
                            writableSheet.addCell(l1);
                        } else if (s1.contains("Mobile")) {
                            Label l1 = new Label(j, i, s1.substring(13, s1.length()));
                            j++;
                            writableSheet.addCell(l1);
                        } else if (s1.contains("VOICE")) {
                            Label l1 = new Label(j, i, s1.substring(10, s1.length()));
                            j++;
                            writableSheet.addCell(l1);
                        } else {
                            Label l1 = new Label(j, i, s1.substring(9, s1.length()));
                            j++;
                            writableSheet.addCell(l1);
                        }
                        s1 = br.readLine();
                        System.out.println(i + "  " + j);
                    }
                    if (!flag) {
                        i++;
                    }
                }
                s = br.readLine();

            }
            writableWorkbook.write();
            writableWorkbook.close();
        } catch (IOException | WriteException e) {
            JOptionPane.showMessageDialog(null, "Exception generated : " + e);
        } finally {
            System.exit(0);
        }

    }

    public static void main(String args[]) {

        java.awt.EventQueue.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    new VCFToExcelSheetConverter().setVisible(true);
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(null, e);
                }
            }
        });
    }
    
    private javax.swing.JButton btnBrowseDestFile;
    private javax.swing.JButton btnBrowseFile;
    private javax.swing.JButton btnStart;
    private javax.swing.JLabel lblDestFile;
    private javax.swing.JLabel lblFileName;
    private javax.swing.JLabel lblSrcFile;
    private javax.swing.JTextField txtDestFile;
    private javax.swing.JTextField txtFileName;
    private javax.swing.JTextField txtSrcFile;
    
}
