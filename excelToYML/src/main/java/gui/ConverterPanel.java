package gui;

import java.awt.Component;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;

import excelToYML.Controller;

public class ConverterPanel extends JPanel implements ActionListener {
	private JButton browseButton, startButton;
	private JTextField textField;
	private JFileChooser chooser;
	private JLabel statusLabel;
	private static final Insets WEST_INSETS = new Insets(1, 0, 1, 5);
	private static final Insets EAST_INSETS = new Insets(1, 5, 1, 0);

	public ConverterPanel() {
		setLayout(new GridBagLayout());
		browseButton = new JButton("Browse");
		browseButton.addActionListener(this);
		startButton = new JButton("Start");
		startButton.setPreferredSize(browseButton.getPreferredSize());
		startButton.addActionListener(this);
		textField = new JTextField(50);
		textField.setEditable(false);
		chooser = new JFileChooser("Choose Excel file") {
			@Override
			public void approveSelection() {
				getSelectedFile().mkdirs();
				super.approveSelection();
			}
		};
		statusLabel = new JLabel("");
		add(browseButton, createGbc(0, 0));
		add(textField, createGbc(1, 0));
		add(startButton, createGbc(0, 1));
		add(statusLabel, createGbc(1, 1));
	}

	@Override
	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == browseButton) {
			if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
				statusLabel.setText("");
				textField.setText(chooser.getSelectedFile().getPath());
			}
		} else if (event.getSource() == startButton) {
			if (chooser.getSelectedFile() != null) {
				statusLabel.setText("");
				Controller.getInstance().readExcelFile(chooser.getSelectedFile().getAbsolutePath());
				String savePath = Controller.getInstance().getSaveFileNameFromPath(chooser.getSelectedFile().getAbsolutePath());
				Controller.getInstance().saveXML(savePath);
				statusLabel.setText("Conversion completed");
			} else {
				JOptionPane.showMessageDialog(null, "Please select file", "Info", JOptionPane.INFORMATION_MESSAGE);
			}
		}
	}

	private void disableEnableAllButtons(JPanel panel, boolean enable) {
		for (Component c : panel.getComponents()) {
			if (c instanceof JButton) {
				((JButton) c).setEnabled(enable);
			}
		}
	}

	private GridBagConstraints createGbc(int x, int y) {
		GridBagConstraints gbc = new GridBagConstraints();
		gbc.gridx = x;
		gbc.gridy = y;
		gbc.gridwidth = 1;
		gbc.gridheight = 1;
		// gbc.anchor = (x == 0) ? GridBagConstraints.WEST : GridBagConstraints.EAST;
		gbc.anchor = (x == 0) ? GridBagConstraints.EAST : GridBagConstraints.WEST;
		// gbc.fill = (x == 0) ? GridBagConstraints.BOTH : GridBagConstraints.HORIZONTAL;
		gbc.insets = (x == 0) ? WEST_INSETS : EAST_INSETS;
		gbc.weightx = (x == 0) ? 0.1 : 1.0;
		gbc.weighty = 1.0;
		return gbc;
	}
}
