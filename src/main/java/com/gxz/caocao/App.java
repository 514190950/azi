package com.gxz.caocao;

import com.gxz.caocao.factory.ButtonFactory;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;


/**
 * @author gxz gongxuanzhang@foxmail.com
 **/
public class App extends Application {


    Button start = new Button("开始游戏");



    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("曹操传");
        BorderPane borderPane = new BorderPane();
        VBox vBox = new VBox();
        vBox.getChildren().addAll(start, ButtonFactory.createExitButton());
        borderPane.setCenter(vBox);
        primaryStage.setScene(new Scene(borderPane));
        primaryStage.setMaximized(true);
        primaryStage.show();

    }


    public static void main(String[] args) {
        launch(args);
    }
}
