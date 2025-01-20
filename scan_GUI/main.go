package main

import (
	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
)

func main() {
    myApp := app.New()
    myWindow := myApp.NewWindow("Nmap扫描工具")
    
    // 创建输入控件
    inputFileBtn := widget.NewButton("选择输入文件", nil)
    inputPathLabel := widget.NewLabel("未选择文件")
    outputNameEntry := widget.NewEntry()
    outputNameEntry.SetPlaceHolder("输出文件名称")
    nmapCmdEntry := widget.NewEntry()
    nmapCmdEntry.SetText("-sV -O")
    
    // 使用固定宽度的容器来控制输入框宽度
    outputContainer := container.NewHBox(
        widget.NewLabel("输出文件名:"),
        container.New(layout.NewMaxLayout(),
            widget.NewLabel(""), // 用于设置最小宽度的占位符
            outputNameEntry,
        ),
    )
    outputContainer.Resize(fyne.NewSize(600, 36))

    cmdContainer := container.NewHBox(
        widget.NewLabel("Nmap命令:"),
        container.New(layout.NewMaxLayout(),
            widget.NewLabel("                                                  "), // 用空格设置最小宽度
            nmapCmdEntry,
        ),
    )
    cmdContainer.Resize(fyne.NewSize(600, 36))
    
    // 创建日志显示区域
    logArea := widget.NewTextGrid()
    logScroll := container.NewScroll(logArea)
    
    // 创建扫描状态显示
    progressBar := widget.NewProgressBar()
    statusLabel := widget.NewLabel("就绪")
    
    // 开始扫描按钮
    startBtn := widget.NewButton("开始扫描", nil)
    startBtn.Disable()
    
    // 设置文件选择按钮回调
    inputFileBtn.OnTapped = func() {
        dialog.ShowFileOpen(func(reader fyne.URIReadCloser, err error) {
            if err != nil {
                dialog.ShowError(err, myWindow)
                return
            }
            if reader == nil {
                return
            }
            inputPathLabel.SetText(reader.URI().Path())
            startBtn.Enable()
        }, myWindow)
    }
    
    // 布局设置
    content := container.NewVBox(
        container.NewHBox(inputFileBtn, inputPathLabel),
        outputContainer,
        cmdContainer,
        startBtn,
        progressBar,
        statusLabel,
        logScroll,
    )
    
    myWindow.SetContent(content)
    myWindow.Resize(fyne.NewSize(800, 600))
    myWindow.ShowAndRun()
} 