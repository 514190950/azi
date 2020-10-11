package com.gxz.easyexcel;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.stage.Stage;

import java.io.File;

/**
 * @author gxz gongxuanzhang@foxmail.com
 **/
public class App extends Application {

    Button button = new Button("gogogo");
    TextField textField = new TextField();

    @Override
    public void start(Stage primaryStage) {
        BorderPane borderPane = new BorderPane();
        HBox hBox = new HBox();
        button.setOnAction((event) -> readExcel());
        hBox.getChildren().addAll(new Label("你excel的全路径："), textField);
        borderPane.setBottom(button);
        borderPane.setCenter(hBox);
        primaryStage.setScene(new Scene(borderPane, 400, 100));
        primaryStage.setTitle("牛逼就完事了");
        primaryStage.show();
    }

    private void readExcel() {
        String filePath = textField.getText();
        File file = new File(filePath);
        if (filePath.endsWith(".xlsx") && file.exists()) {
            ExcelAnalysis.analysis(file);
        } else {
            alert("认真点啊，这个路径找不到文件");
        }
    }

    private void alert(String message) {
        Alert alert = new Alert(Alert.AlertType.ERROR);
        alert.setTitle("行不行啊");
        alert.setContentText(message);
        alert.showAndWait();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
