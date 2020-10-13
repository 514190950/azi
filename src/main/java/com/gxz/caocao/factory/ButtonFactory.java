package com.gxz.caocao.factory;

import javafx.scene.control.Button;

/**
 * @author gxz gongxuanzhang@foxmail.com
 *
 **/
public class ButtonFactory {
    public static Button createExitButton(){
        Button button = new Button("退出游戏");
        button.setOnAction((e)->System.exit(0));
        return button;
    }
}
