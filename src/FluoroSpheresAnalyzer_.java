import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumnModel;
import javax.swing.table.TableModel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import ij.IJ;
import ij.ImagePlus;
import ij.ImageStack;
import ij.gui.Overlay;
import ij.gui.Roi;
import ij.gui.ShapeRoi;
import ij.plugin.ChannelSplitter;
import ij.plugin.PlugIn;
import ij.plugin.ZProjector;
import ij.plugin.frame.RoiManager;
import ij.process.ImageProcessor;
import loci.plugins.in.DisplayHandler;
import loci.plugins.in.ImportProcess;
import loci.plugins.in.ImporterOptions;

public class FluoroSpheresAnalyzer_ implements PlugIn {
	ImagePlus impMaxProjection;
	ImagePlus[] impMaxProjectionSlices, impMaxProjectionSlicesMeasure, impsSlices, channels, channelsDup;
	Roi[] selectionRois, selectionRoisChannel;
	RoiManager rm;
	ArrayList<List<Roi>> listOfLists = new ArrayList<List<Roi>>();
	JButton exportButton;
	Object[] columnNames = new Object[] { "", "", "" }, columnNamesDef = new Object[] { "Ch1", "Ch2", "Ch3", "Ch4" };
	JScrollPane jScrollPaneDef;
	List<String> meanValues = new ArrayList<String>();
	JTable tableDef;

	@Override
	public void run(String arg0) {
		ImagePlus imp = IJ.getImage();
		IJ.run("ROI Manager...", "");
		rm = RoiManager.getInstance();
		if (rm != null)
			rm.reset();

		channels = ChannelSplitter.split(imp);
		channelsDup = new ImagePlus[channels.length];
		impMaxProjection = ZProjector.run(imp, "max");
		impMaxProjectionSlicesMeasure = ChannelSplitter.split(impMaxProjection.duplicate());
		impMaxProjectionSlices = ChannelSplitter.split(impMaxProjection);
		selectionRois = new Roi[impMaxProjectionSlices.length];
		selectionRoisChannel = new Roi[channels.length];
		// IJ.log(" ");
		// IJ.log("-Image Analyzed: ");
		// IJ.log(" " + imp.getTitle());
		// IJ.log(" ");
		if (imp.getNSlices() == 1)
		for (int j = 0; j < channels.length; j++) {
			// IJ.log("pasa");
			channelsDup[j] = channels[j].duplicate();
		
			IJ.run(channelsDup[j], "Auto Threshold", "method=Otsu ignore_black white");
			IJ.run(channelsDup[j], "Create Selection", "");
			selectionRoisChannel[j] = channelsDup[j].getRoi();
			rm.addRoi(selectionRoisChannel[j]);
			channels[j].setRoi(selectionRoisChannel[j]);
			meanValues.add(String.valueOf(selectionRoisChannel[j].getStatistics().mean));
			// IJ.log(" " + "-Label: " + channels[j].getTitle() + ": "
			// + selectionRoisChannel[j].getName() + " -Area: "
			// + selectionRoisChannel[j].getStatistics().area + " -Mean: "
			// + selectionRoisChannel[j].getStatistics().mean + " -Channel: " + (j + 1));
			// IJ.log(" ");
		}
		if (imp.getNSlices() != 1) {
			for (int i = 0; i < impMaxProjectionSlices.length; i++) {

				IJ.run(impMaxProjectionSlices[i], "Auto Threshold", "method=Otsu ignore_black white");
				// IJ.run(impMaxProjectionSlices[i], "Watershed", "");
				IJ.run(impMaxProjectionSlices[i], "Create Selection", "");
				selectionRois[i] = impMaxProjectionSlices[i].getRoi();
				rm.addRoi(selectionRois[i]);
				impMaxProjectionSlicesMeasure[i].setRoi(selectionRois[i]);
				meanValues.add(String.valueOf(selectionRois[i].getStatistics().mean));
				// IJ.log(" " + "-Label: " + impMaxProjectionSlices[i].getTitle() + ": "
				// + selectionRois[i].getName() + " -Area: " +
				// selectionRois[i].getStatistics().area
				// + " -Mean: " + selectionRois[i].getStatistics().mean + " -Channel: " + (i +
				// 1));
				// IJ.log(" ");
				// rm.runCommand(impMaxProjectionSlicesMeasure[i], "Measure");

			}
		}
		processTable(meanValues);

	}

	public static ImagePlus[] stack2images(ImagePlus imp) {
		String sLabel = imp.getTitle();
		String sImLabel = "";
		ImageStack stack = imp.getStack();

		int sz = stack.getSize();
		int currentSlice = imp.getCurrentSlice(); // to reset ***

		DecimalFormat df = new DecimalFormat("0000"); // for title
		ImagePlus[] arrayOfImages = new ImagePlus[imp.getStack().getSize()];
		for (int n = 1; n <= sz; ++n) {
			imp.setSlice(n); // activate next slice ***

			// Get current image processor from stack. What ever is
			// used here should do a COPY pixels from old processor to
			// new. For instance, ImageProcessor.crop() returns copy.
			ImageProcessor ip = imp.getProcessor(); // ***
			ImageProcessor newip = ip.createProcessor(ip.getWidth(), ip.getHeight());
			newip.setPixels(ip.getPixelsCopy());

			// Create a suitable label, using the slice label if possible
			sImLabel = imp.getStack().getSliceLabel(n);
			if (sImLabel == null || sImLabel.length() < 1) {
				sImLabel = "slice" + df.format(n) + "_" + sLabel;
			}
			// Create new image corresponding to this slice.
			ImagePlus im = new ImagePlus(sImLabel, newip);
			im.setCalibration(imp.getCalibration());
			arrayOfImages[n - 1] = im;

			// Show this image.
			// imp.show();
		}
		// Reset original stack state.
		imp.setSlice(currentSlice); // ***
		if (imp.isProcessor()) {
			ImageProcessor ip = imp.getProcessor();
			ip.setPixels(ip.getPixels()); // ***
		}
		imp.setSlice(currentSlice);
		return arrayOfImages;
	}

	public void processTable(List<String> meanValue) {

		exportButton = new JButton("Export Table");

		tableDef = new JTable();
		DefaultTableModel modelDef = new DefaultTableModel();
		modelDef.setColumnIdentifiers(columnNamesDef);
		jScrollPaneDef = new JScrollPane(tableDef);
		jScrollPaneDef.setPreferredSize(new Dimension(650, 300));
		Object[][] dataTImages = new Object[1][columnNamesDef.length];
		for (int i = 0; i < dataTImages.length; i++)
			for (int j = 0; j < dataTImages[i].length; j++)
				dataTImages[i][j] = "";
		modelDef = new DefaultTableModel(dataTImages, columnNamesDef) {

			@Override
			public Class<?> getColumnClass(int column) {
				if (getRowCount() >= 0) {
					Object value = getValueAt(0, column);
					if (value != null) {
						return getValueAt(0, column).getClass();
					}
				}

				return super.getColumnClass(column);
			}

			public boolean isCellEditable(int row, int col) {
				return false;
			}

		};
		tableDef.setModel(modelDef);
		for (int i = 0; i < modelDef.getRowCount(); i++) {
			modelDef.setValueAt(meanValues.get(0), i, tableDef.convertColumnIndexToModel(0));
			modelDef.setValueAt(meanValues.get(1), i, tableDef.convertColumnIndexToModel(1));
			modelDef.setValueAt(meanValues.get(2), i, tableDef.convertColumnIndexToModel(2));
			modelDef.setValueAt(meanValues.get(3), i, tableDef.convertColumnIndexToModel(3));

		}

		tableDef.setModel(modelDef);
		tableDef.setSelectionBackground(new Color(229, 255, 204));
		tableDef.setSelectionForeground(new Color(0, 102, 0));
		DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
		centerRenderer.setHorizontalAlignment(JLabel.CENTER);
		tableDef.setDefaultRenderer(String.class, centerRenderer);
		tableDef.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		tableDef.setRowHeight(60);
		tableDef.setAutoCreateRowSorter(true);
		for (int u = 0; u < tableDef.getColumnCount(); u++)
			tableDef.getColumnModel().getColumn(u).setPreferredWidth(170);
		tableDef.getColumnModel().getColumn(0).setPreferredWidth(250);

		JPanel mainPanel = new JPanel();
		mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));
		JPanel imagePanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
		imagePanel.add(jScrollPaneDef);
		JPanel exportPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
		exportPanel.add(exportButton);
		mainPanel.add(imagePanel);
		mainPanel.add(exportPanel);
		JFrame frameDef = new JFrame();
		frameDef.setTitle("Results");
		frameDef.setResizable(false);
		frameDef.add(mainPanel);
		frameDef.pack();
		frameDef.setSize(660, 400);
		frameDef.setLocationRelativeTo(null);
		frameDef.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		frameDef.setVisible(true);
		exportButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {

				try {
					csvExport(tableDef);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}

			}
		});

	}

	public void csvExport(JTable tableDef) throws IOException {

		try (HSSFWorkbook fWorkbook = new HSSFWorkbook()) {
			HSSFSheet fSheet = fWorkbook.createSheet("new Sheet");
			HSSFFont sheetTitleFont = fWorkbook.createFont();
			HSSFCellStyle cellStyle = fWorkbook.createCellStyle();
			sheetTitleFont.setBold(true);
			// sheetTitleFont.setColor();
			TableModel model = tableDef.getModel();

			// Get Header
			TableColumnModel tcm = tableDef.getColumnModel();
			HSSFRow hRow = fSheet.createRow((short) 0);
			for (int j = 0; j < tcm.getColumnCount(); j++) {
				HSSFCell cell = hRow.createCell((short) j);
				cell.setCellValue(tcm.getColumn(j).getHeaderValue().toString());
				cell.setCellStyle(cellStyle);
			}

			// Get Other details
			for (int i = 0; i < model.getRowCount(); i++) {
				HSSFRow fRow = fSheet.createRow((short) i + 1);
				for (int j = 0; j < model.getColumnCount(); j++) {
					HSSFCell cell = fRow.createCell((short) j);
					cell.setCellValue(model.getValueAt(i, j).toString());
					cell.setCellStyle(cellStyle);
				}
			}
			JFrame parentFrame = new JFrame();

			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			fileChooser.setDialogTitle("Specify a path to save");
			int userSelection = fileChooser.showSaveDialog(parentFrame);
			fileChooser.setSelectedFile(new File("FluoroSpheres_data.xlsx"));
			String pathSave = null;
			if (userSelection == JFileChooser.APPROVE_OPTION) {
				File fileToSave = fileChooser.getCurrentDirectory();
				pathSave = fileToSave.getAbsolutePath();
			}
			FileOutputStream fileOutputStream;
			fileOutputStream = new FileOutputStream(pathSave + File.separator + "FluoroSpheres_data.xlsx");
			try (BufferedOutputStream bos = new BufferedOutputStream(fileOutputStream)) {
				fWorkbook.write(bos);
			}
			fileOutputStream.close();

			JOptionPane.showMessageDialog(null, "FluoroSpheres_data.xlsx exported to " + pathSave);
		}

	}

}
