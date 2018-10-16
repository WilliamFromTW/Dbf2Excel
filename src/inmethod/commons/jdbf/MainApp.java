package inmethod.commons.jdbf;

import java.io.File;

import javax.swing.JDialog;
import javax.swing.JOptionPane;

import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.FileDialog;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.custom.CLabel;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.events.MouseAdapter;
import org.eclipse.swt.events.MouseEvent;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;

public class MainApp {

	protected Shell shell;
	private Text txtBig;
	private Text textInputDbf;
	private Text textOutputExcel;

	/**
	 * Launch the application.
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		try {
			MainApp window = new MainApp();
			window.open();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Open the window.
	 */
	public void open() {
		Display display = Display.getDefault();
		createContents();
		shell.open();
		shell.layout();
		while (!shell.isDisposed()) {
			if (!display.readAndDispatch()) {
				display.sleep();
			}
		}
	}

	/**
	 * Create contents of the window.
	 */
	protected void createContents() {
		shell = new Shell();
		shell.setSize(450, 211);
		shell.setText("InMethodJdbf");

		Button btnExporttoexcel = new Button(shell, SWT.NONE);
		btnExporttoexcel.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseUp(MouseEvent e) {
				if (txtBig.getText().trim().equals("") || textInputDbf.getText().trim().equals("")
						|| textOutputExcel.getText().trim().equals("")) {
					JOptionPane.showMessageDialog(null, "DBF file or Output directory is required!", "Required",
							JOptionPane.WARNING_MESSAGE);
				}else {
					try {
						JOptionPane pane = new JOptionPane("Please wait, dialog will auto close ");
						pane.setOptions(new Object[] {});
						JDialog dialog = pane.createDialog("Processing command");
						dialog.setModal(false);
						dialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
						dialog.setVisible(true);
						//System.out.println(f.getAbsolutePath().substring(f.getAbsolutePath().lastIndexOf("\\")+1));
						String sFileName = textOutputExcel.getText().trim()+File.separator+textInputDbf.getText().trim().substring(textInputDbf.getText().trim().lastIndexOf(File.separator)+1);
						sFileName = sFileName.substring(0,sFileName.lastIndexOf("."))+".xlsx";
						//System.out.println(sFileName);
						JdbfUtilty.getInstance().exportDbfToExcel(txtBig.getText().trim(),textInputDbf.getText().trim() , sFileName);
						JOptionPane.showMessageDialog(null, "Export DBF to excel successfully", "Status",
								JOptionPane.INFORMATION_MESSAGE);
						dialog.dispose();
					}catch(Exception ee) {
						JOptionPane.showMessageDialog(null, ee.getMessage(), "ERROR",
								JOptionPane.ERROR_MESSAGE);
					}
				}
			}
		});
		btnExporttoexcel.addSelectionListener(new SelectionAdapter() {
			@Override
			public void widgetSelected(SelectionEvent e) {
			}
		});
		btnExporttoexcel.setBounds(331, 140, 93, 25);
		btnExporttoexcel.setText("ExportToExcel");

		CLabel lblCharset = new CLabel(shell, SWT.NONE);
		lblCharset.setToolTipText("BIG5, UTF8, UTF-8 , CP1250");
		lblCharset.setBounds(10, 29, 66, 21);
		lblCharset.setText("CharSet");

		txtBig = new Text(shell, SWT.BORDER | SWT.CENTER);
		txtBig.setText("big5");
		txtBig.setBounds(107, 29, 73, 21);

		CLabel lblDbfFile = new CLabel(shell, SWT.NONE);
		lblDbfFile.setBounds(10, 62, 66, 21);
		lblDbfFile.setText("DBF File");

		CLabel lblOutputExcel = new CLabel(shell, SWT.NONE);
		lblOutputExcel.setBounds(10, 93, 83, 21);
		lblOutputExcel.setText("Output Excel");
		FileDialog inputFD = new FileDialog(shell, SWT.SINGLE);

		DirectoryDialog outputExcel = new DirectoryDialog(shell);

		textInputDbf = new Text(shell, SWT.BORDER);
		textInputDbf.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseUp(MouseEvent e) {

				if (inputFD.open() != null) {
					String[] names = inputFD.getFileNames();
					for (int i = 0, n = names.length; i < n; i++) {
						StringBuffer buf = new StringBuffer(inputFD.getFilterPath());
						if (buf.charAt(buf.length() - 1) != File.separatorChar)
							buf.append(File.separatorChar);
						buf.append(names[i]);
						textInputDbf.setText(buf.toString());

					}
				}
			}
		});
		textInputDbf.setBounds(107, 62, 317, 21);

		textOutputExcel = new Text(shell, SWT.BORDER);
		textOutputExcel.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseUp(MouseEvent e) {
				outputExcel.setFilterPath("" + File.separatorChar); // Windows specific
				textOutputExcel.setText(outputExcel.open());
			}

		});
		textOutputExcel.setBounds(107, 93, 317, 21);

	}
}
