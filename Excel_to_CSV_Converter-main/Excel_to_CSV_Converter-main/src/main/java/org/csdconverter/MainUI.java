package org.csdconverter;

import java.io.File;
import java.util.List;

import javafx.application.Application;
import javafx.beans.property.SimpleStringProperty;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableView;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class MainUI extends Application {

    private TableView<SheetConfigTableModel> tableView;

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel to CSV Converter");

        VBox layout = new VBox(10);
        layout.setPadding(new Insets(20, 20, 20, 20));

        FileChooser configFileChooser = new FileChooser();
        configFileChooser.setTitle("Select Configuration File");

        Button configFileButton = new Button("Select Configuration File");
        Label configFilePathLabel = new Label("No file selected");

        configFileButton.setOnAction(e -> {
            File configFile = configFileChooser.showOpenDialog(primaryStage);
            if (configFile != null) {
                configFilePathLabel.setText(configFile.getAbsolutePath());
                loadSheetConfigs(configFile.getAbsolutePath());
            }
        });

        FileChooser excelFileChooser = new FileChooser();
        excelFileChooser.setTitle("Select Excel File");

        Button excelFileButton = new Button("Select Excel File");
        Label excelFilePathLabel = new Label("No file selected");

        excelFileButton.setOnAction(e -> {
            File excelFile = excelFileChooser.showOpenDialog(primaryStage);
            if (excelFile != null) {
                excelFilePathLabel.setText(excelFile.getAbsolutePath());
            }
        });

        tableView = new TableView<>();
        tableView.setColumnResizePolicy(TableView.CONSTRAINED_RESIZE_POLICY);
        TableColumn<SheetConfigTableModel, String> sheetNameColumn = new TableColumn<>("Sheet Name");
        sheetNameColumn.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().getSheetName()));

        TableColumn<SheetConfigTableModel, String> csvNameColumn = new TableColumn<>("CSV Name");
        csvNameColumn.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().getCsvName()));

        TableColumn<SheetConfigTableModel, String> transposeColumn = new TableColumn<>("Transpose");
        transposeColumn.setCellValueFactory(data -> new SimpleStringProperty(data.getValue().isTranspose()));

        tableView.getColumns().addAll(sheetNameColumn, csvNameColumn, transposeColumn);

        Button startButton = new Button("Start Conversion");
        startButton.setOnAction(e -> {
            String configFilePath = configFilePathLabel.getText();
            String excelFilePath = excelFilePathLabel.getText();
            if (!configFilePath.equals("No file selected") && !excelFilePath.equals("No file selected")) {
                MainCSD.convert(configFilePath, excelFilePath);
                showAlert(Alert.AlertType.INFORMATION, "Conversion Complete", "The conversion process has completed successfully.");
            } else {
                showAlert(Alert.AlertType.WARNING, "Files Missing", "Please select both configuration and Excel files.");
            }
        });

        layout.getChildren().addAll(
                configFileButton, configFilePathLabel,
                excelFileButton, excelFilePathLabel,
                tableView,
                startButton
        );

        Scene scene = new Scene(layout, 800, 600);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void loadSheetConfigs(String configFilePath) {
        List<SheetConfig> configs = MainCSD.loadSheetConfigs(configFilePath);
        ObservableList<SheetConfigTableModel> sheetConfigs = FXCollections.observableArrayList();
        for (SheetConfig config : configs) {
            sheetConfigs.add(new SheetConfigTableModel(config));
        }
        tableView.setItems(sheetConfigs);
    }

    private void showAlert(Alert.AlertType type, String title, String message) {
        Alert alert = new Alert(type);
        alert.setTitle(title);
        alert.setHeaderText(null);
        alert.setContentText(message);
        alert.showAndWait();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
