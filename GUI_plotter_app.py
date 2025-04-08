import sys
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QComboBox,
    QFileDialog, QHBoxLayout, QLineEdit, QRadioButton, QButtonGroup,
    QCheckBox, QListWidget, QListWidgetItem, QMessageBox
)
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas


class PlotApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced Excel/CSV Plotter")
        self.setGeometry(100, 100, 1000, 800)
        self.data = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        self.load_button = QPushButton("Load Excel/CSV/TSV File")
        self.load_button.clicked.connect(self.load_file)
        layout.addWidget(self.load_button)

        self.x_list = QListWidget()
        self.x_list.setSelectionMode(QListWidget.MultiSelection)
        self.y_list = QListWidget()
        self.y_list.setSelectionMode(QListWidget.MultiSelection)

        hlayout1 = QHBoxLayout()
        hlayout1.addWidget(QLabel("X-axis (max 10):"))
        hlayout1.addWidget(self.x_list)
        hlayout1.addWidget(QLabel("Y-axis (max 10):"))
        hlayout1.addWidget(self.y_list)
        layout.addLayout(hlayout1)

        self.color_dropdowns = []
        for i in range(10):
            cb = QComboBox()
            cb.addItems(sns.color_palette().as_hex())
            cb.setCurrentIndex(i % len(sns.color_palette()))
            self.color_dropdowns.append(cb)
            layout.addWidget(cb)

        self.plot_group = QButtonGroup(self)
        self.plot_types = ['Boxplot', 'Scatter', 'Line', 'Histogram', 'Bar', 'Violin']
        type_layout = QHBoxLayout()
        for ptype in self.plot_types:
            btn = QRadioButton(ptype)
            self.plot_group.addButton(btn)
            type_layout.addWidget(btn)
        self.plot_group.buttons()[0].setChecked(True)
        layout.addLayout(type_layout)

        self.title_input = QLineEdit()
        self.title_input.setPlaceholderText("Plot Title")
        self.xlabel_input = QLineEdit()
        self.xlabel_input.setPlaceholderText("X-axis Label")
        self.ylabel_input = QLineEdit()
        self.ylabel_input.setPlaceholderText("Y-axis Label")
        layout.addWidget(self.title_input)
        layout.addWidget(self.xlabel_input)
        layout.addWidget(self.ylabel_input)

        self.grid_checkbox = QCheckBox("Show Grid")
        layout.addWidget(self.grid_checkbox)

        self.save_input = QLineEdit()
        self.save_input.setPlaceholderText("Filename to save (without extension)")
        layout.addWidget(self.save_input)

        self.plot_button = QPushButton("Generate Plot")
        self.plot_button.clicked.connect(self.generate_plot)
        layout.addWidget(self.plot_button)

        self.export_plotly_btn = QPushButton("Export as Interactive Plot")
        self.export_plotly_btn.clicked.connect(self.export_plotly_plot)
        layout.addWidget(self.export_plotly_btn)

        self.canvas = FigureCanvas(plt.figure())
        layout.addWidget(self.canvas)

        self.stats_label = QLabel("Stats will appear here.")
        layout.addWidget(self.stats_label)

        self.setLayout(layout)

    def load_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Open File", "", "Data Files (*.xlsx *.csv *.tsv)")
        if file_name:
            try:
                if file_name.endswith(".xlsx"):
                    self.data = pd.read_excel(file_name)
                elif file_name.endswith(".csv"):
                    self.data = pd.read_csv(file_name)
                elif file_name.endswith(".tsv"):
                    self.data = pd.read_csv(file_name, sep='\t')
                else:
                    raise ValueError("Unsupported file format.")

                self.x_list.clear()
                self.y_list.clear()
                for col in self.data.columns:
                    self.x_list.addItem(QListWidgetItem(col))
                    self.y_list.addItem(QListWidgetItem(col))

            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def generate_plot(self):
        if self.data is None:
            return

        x_cols = [item.text() for item in self.x_list.selectedItems()]
        y_cols = [item.text() for item in self.y_list.selectedItems()]
        if len(x_cols) > 10 or len(y_cols) > 10:
            self.stats_label.setText("Select up to 10 columns for both X and Y.")
            return

        title = self.title_input.text()
        xlabel = self.xlabel_input.text()
        ylabel = self.ylabel_input.text()
        show_grid = self.grid_checkbox.isChecked()
        save_name = self.save_input.text()
        plot_type = self.plot_group.checkedButton().text().lower()

        self.canvas.figure.clf()
        ax = self.canvas.figure.add_subplot(111)

        try:
            for i, y in enumerate(y_cols):
                color = self.color_dropdowns[i % len(self.color_dropdowns)].currentText()
                x = x_cols[i % len(x_cols)] if x_cols else None

                if plot_type == "boxplot":
                    sns.boxplot(data=self.data, x=x, y=y, ax=ax, color=color)
                elif plot_type == "violin":
                    sns.violinplot(data=self.data, x=x, y=y, ax=ax, color=color)
                elif plot_type == "scatter":
                    ax.scatter(self.data[x], self.data[y], color=color)
                elif plot_type == "line":
                    ax.plot(self.data[x], self.data[y], color=color, label=y)
                elif plot_type == "histogram":
                    ax.hist(self.data[y], color=color, alpha=0.5)
                elif plot_type == "bar":
                    ax.bar(self.data[x], self.data[y], color=color)

            ax.set_title(title)
            ax.set_xlabel(xlabel)
            ax.set_ylabel(ylabel)
            ax.grid(show_grid)
            if plot_type in ["line"]:
                ax.legend()
            self.canvas.draw()

            if save_name:
                self.canvas.figure.savefig(f"{save_name}.png")

            if y_cols and pd.api.types.is_numeric_dtype(self.data[y_cols[0]]):
                stats = self.data[y_cols[0]].describe()
                self.stats_label.setText(f"Mean: {stats['mean']:.2f}, Median: {self.data[y_cols[0]].median():.2f}, Std: {stats['std']:.2f}")
            else:
                self.stats_label.setText("Y-axis must be numeric for stats.")

        except Exception as e:
            self.stats_label.setText(f"Error: {str(e)}")

    def export_plotly_plot(self):
        if self.data is None:
            return

        x_cols = [item.text() for item in self.x_list.selectedItems()]
        y_cols = [item.text() for item in self.y_list.selectedItems()]
        plot_type = self.plot_group.checkedButton().text().lower()

        try:
            fig = None
            color = self.color_dropdowns[0].currentText()
            x = x_cols[0] if x_cols else None
            y = y_cols[0] if y_cols else None

            if plot_type == 'scatter':
                fig = px.scatter(self.data, x=x, y=y, color_discrete_sequence=[color])
            elif plot_type == 'line':
                fig = px.line(self.data, x=x, y=y, color_discrete_sequence=[color])
            elif plot_type == 'bar':
                fig = px.bar(self.data, x=x, y=y, color_discrete_sequence=[color])
            elif plot_type == 'histogram':
                fig = px.histogram(self.data, x=y, color_discrete_sequence=[color])
            elif plot_type == 'boxplot':
                fig = px.box(self.data, x=x, y=y, color_discrete_sequence=[color])
            elif plot_type == 'violin':
                fig = px.violin(self.data, x=x, y=y, color_discrete_sequence=[color])
            else:
                self.stats_label.setText("Plotly does not support this plot type.")
                return

            fig.update_layout(title=self.title_input.text())
            save_path, _ = QFileDialog.getSaveFileName(self, "Save HTML", "", "HTML Files (*.html)")
            if save_path:
                fig.write_html(save_path)
                self.stats_label.setText(f"Interactive plot saved as {save_path}")

        except Exception as e:
            self.stats_label.setText(f"Plotly Export Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PlotApp()
    window.show()
    sys.exit(app.exec_())


