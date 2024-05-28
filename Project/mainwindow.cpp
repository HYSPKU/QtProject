#include "mainwindow.h"
#include "./ui_mainwindow.h"
#include <QFileDialog>
#include<QAction>
#include<QAbstractItemView>
#include<QMenu>
#include<QTime>
#include<QLineEdit>
#include<vector>
#include<QMediaPlayer>
#include <QMediaContent>
#include <QUrl>
#include<QInputDialog>
#include<QMessageBox>
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent),ui(new Ui::MainWindow),timer(new QTimer(this))
{
    ui->setupUi(this);
    connect(timer,&QTimer::timeout,this,&MainWindow::updateTimeDisplay);
    timer->start(1000);
    ui->timeDisplay->setReadOnly(true);
    updateTimeDisplay();
    connect(ui->_calendar, &QCalendarWidget::selectionChanged, this, &MainWindow::handleSelectionChanged);
    connect(ui->_buttonConfirm,&QPushButton::clicked,this,&MainWindow::Submit);
    connect(ui->_buttonInport,&QPushButton::clicked,this,&MainWindow::ClassImport);
    connect(ui->_buttonModify,&QPushButton::clicked,this,&MainWindow::ClassModify);
    connect(ui->_buttonPersonalize,&QPushButton::clicked,this,&MainWindow::Personalize);
    connect(ui->_calendar,&QCalendarWidget::activated,this,&MainWindow::oneday);
    ui->_table->setContextMenuPolicy(Qt::CustomContextMenu);
    connect(ui->_table, &QTableWidget::customContextMenuRequested, this, &MainWindow::onItemContextMenuRequested);
    AddRow(0);
    for(int i = 0; i < 5; ++ i){
        AddColumn(i);
    }
    AddItem(0,0,"起始时间");
    AddItem(0,1,"终止时间");
    AddItem(0,2,"事件");
    AddItem(0,3,"地点");
    AddItem(0,4,"备注");
    files.getNodes(QString("../Project/nodes.csv"));

    qDebug() << "修改之前";
    qDebug() << files.readUserInfo(0);
    qDebug() << files.readUserInfo(1);
    files.saveUserInfo(1);
    files.saveUserInfo(1);
    qDebug() << "修改之后";
    qDebug() << files.readUserInfo(0);
    qDebug() << files.readUserInfo(1);


    configs = new Config(this);

}

MainWindow::~MainWindow()
{
}

void MainWindow::updateTimeDisplay(){
    QTime currentTime=QTime::currentTime();
    QString timeString;
    int idx = -1;
    for(int i = activities.size() - 1; i >= 0; -- i){
        if(activities[i].begin < currentTime){
            idx = i;
            break;
        }
    }
    if(idx!=-1&&idx!=activities.size()-1){
        QTime next=activities[idx+1].begin;
        qint64 msecs = currentTime.msecsTo(next);
        // 将毫秒转换为小时、分钟和秒
        int hours = msecs / (1000 * 60 * 60);
        int minutes = (msecs % (1000 * 60 * 60)) / (1000 * 60);
        int seconds = (msecs % (1000 * 60)) / 1000;
        // 格式化时间差
        timeString = QString("%1:%2:%3").arg(hours, 2, 10, QChar('0')).arg(minutes, 2, 10, QChar('0')).arg(seconds, 2, 10, QChar('0'));
    }
    else if(activities.size()==0){
        timeString="日程未规划";
    }
    else{
        timeString="今日任务已完成";
    }
    // 设置 QLineEdit 的文本
    ui->timeDisplay->setText(timeString);
    if(timeString=="00:00:01"){
       QMediaPlayer mediaPlayer;
       mediaPlayer.setMedia(QMediaContent(QUrl::fromLocalFile("../Project/bell.wav")));
           qDebug() << "Played!";
           mediaPlayer.setVolume(50);
           mediaPlayer.play(); // 播放声音
    }
}


void MainWindow::ClassImport(){
    QString fileName = QFileDialog::getOpenFileName(this,
                                                        tr("打开文件"), "", tr("Excel 文件 (*.xls *.xlsx)")); // 只允许选择 .xls 或 .xlsx 文件
    if (fileName.isEmpty()) { // 如果没有选择文件，直接返回
            return;
    }
    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);

    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)", fileName);
    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目
    if(sheet_count > 0)
    {
        QAxObject *work_sheet = work_book->querySubObject("Sheets(int)", 1);
        QAxObject *used_range = work_sheet->querySubObject("UsedRange");
        QAxObject *rows = used_range->querySubObject("Rows");
        int row_count = rows->property("Count").toInt();  //获取行数
        for(int i=2;i<=8;i++){
            int j=2;
            while(j<=13){
                QString txtt = work_sheet->querySubObject("Cells(int,int)",j,i)->dynamicCall("Value2()").toString(); //获取单元格内容
                if(txtt=="" ){
                    j++;
                    continue;
                }
                int leftbracket=0,rightbracket=0;
                int len=txtt.length();
                int f=0;
                Event t;

                for(int k=0;k<len;k++){
                    if(txtt[k]=='(') leftbracket=k;

                    if(txtt[k]==')') {
                        f=0;
                        rightbracket=k;
                        t.Sposition=txtt.mid(leftbracket+1,rightbracket-leftbracket-1);
                        for(auto u = files.nameTint.begin();u!=files.nameTint.end();u++){
        //                    qDebug() << u.key() << u.value();
                            if(t.Sposition.contains(u.key())){
                                t.Sposition=u.key();
                                t.iposition=u.value();
                                f=1;
                                break;
                            }
                        }
                        if(f==0){
                            leftbracket=0;
                            continue;
                        }
                        rightbracket=k;
                        break;
                    }
                }
                //qDebug()<<leftbracket<<' '<<rightbracket<<endl;

                t.Sname=txtt.left(leftbracket);
                t.begin=classstart[j-2];
                t.end=classend[j-2];
                t.dayidx=i-1;
                while(j<=13 and work_sheet->querySubObject("Cells(int,int)",j,i)->dynamicCall("Value2()").toString()==txtt){
                    t.end=classend[j-2];
                    j++;
                }
                classschedule.week[i-2].push_back(t);
            }
        }


        work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
        excel.dynamicCall("Quit(void)");  //退出
        return ;
    }

}
void MainWindow::oneday(const QDate&date){
    if(Memo.find(date) != Memo.end()){
        tableclear();
        activities.clear();
        for(Event x: Memo[date]){
            activities.push_back(x);
            AddEvent(x);
        }
        return;
    }

    int dayidx=(date.dayOfWeek()-2)%7+1;
    dayindex=(date.dayOfWeek()-2)%7+1;

    MainWindow::tableclear();

    activities.clear();
    for(Event x:classschedule.week[dayidx]){
        InsertEvent(x);;
    }

    if(configs->Mode() == 2){
        Bailan();
        return;
    }
    GetStudy();

    GetFood();

    GetActivity();

    Memo.insert(date,activities);
}
void MainWindow::tableclear(){
    int t=ui->_table->rowCount();
    for(int i=t-1;i>0;i--){
        ui->_table->removeRow(i);
    }
}



QWidget*& MainWindow::pkumap(){
    return _pkumap;
}

void MainWindow::ShowMap(){
    _pkumap->show();
}
void MainWindow::handleSelectionChanged(){}

void MainWindow::ClassModify() {
    QDate d = ui->_calendar->selectedDate();
    Memo.erase(Memo.find(d));
    oneday(d);
}

void MainWindow::onItemContextMenuRequested(const QPoint& pos) {
    QTableWidgetItem* selectedItem = ui->_table->itemAt(pos);
    if (selectedItem) {
        QMenu* menu = new QMenu(ui->_table);
        QAction* actiondelete = new QAction(tr("删除"), this);
        actiondelete->setData(1);
        menu->addAction(actiondelete);
        connect(actiondelete, &QAction::triggered, [this, selectedItem]() {
            onActiondeleteTriggered(selectedItem);
        });
        QAction* actionadd = menu->addAction("添加");
        connect(actionadd, &QAction::triggered, this,
            onActionaddTriggered);
        QAction* actioncancel = menu->addAction("取消");
        connect(actioncancel, &QAction::triggered, [this, selectedItem]() {
            onActioncancelTriggered(selectedItem);
        });
        QAction* actionrevise = menu->addAction("修改");
        connect(actionrevise, &QAction::triggered, [this, selectedItem,pos]() {
            onActionreviseTriggered(selectedItem);
        });
        // 添加更多菜单项...

        // 显示右键菜单
        menu->exec(QCursor::pos());
    }
}

void MainWindow::onActionreviseTriggered(QTableWidgetItem *item){
    int row=item->row();
    int col=item->column();
    if(col==0){
        QDialog *dialogrevise=new QDialog(this);
        QHBoxLayout* layout = new QHBoxLayout(dialogrevise);
        QPushButton* saveButton=new QPushButton("Save");
        layout->addWidget(saveButton);
        QTimeEdit* timeEdit=new QTimeEdit;
        layout->addWidget(timeEdit);
        connect(saveButton, &QPushButton::clicked, this, [=]() {
            (activities.begin()+row-1)->begin=timeEdit->time();
            dialogrevise->close();
        });
        dialogrevise->exec();
        SortEvent();
        }
    if(col==1){
        QDialog *dialogrevise=new QDialog(this);
        QHBoxLayout* layout = new QHBoxLayout(dialogrevise);
        QPushButton* saveButton=new QPushButton("Save");
        QTimeEdit* timeEdit=new QTimeEdit;
        layout->addWidget(saveButton);
        layout->addWidget(timeEdit);
        connect(saveButton, &QPushButton::clicked, this, [=]() {
            (activities.begin()+row-1)->end=timeEdit->time();
            dialogrevise->close();
        });
        dialogrevise->exec();
        SortEvent();
        }
    if(col==2){
        bool ok;
        QString text = QInputDialog::getText(nullptr, "Enter Text", "Please enter some text:", QLineEdit::Normal, "", &ok);

            if (ok && !text.isEmpty()) {
                activities[row-1].Sname=text;
                SortEvent();
            }
    }
    if(col==3){
        QComboBox* locationEdit = new QComboBox(ui->_table);
        for(auto nm:files.nameTint.keys()){
            locationEdit->addItem(nm);
        }
        QDialog *dialogrevise=new QDialog(this);
        QVBoxLayout* layout = new QVBoxLayout(dialogrevise);
        QPushButton* saveButton=new QPushButton("Save");
        layout->addWidget(saveButton);
        layout->addWidget(locationEdit);
        connect(saveButton, &QPushButton::clicked, this, [=]() {
            (activities.begin()+row-1)->Sposition=locationEdit->currentText();
            dialogrevise->close();
        });
        dialogrevise->exec();
        activities[row-1].iposition=files.nameTint[activities[row-1].Sposition];
        SortEvent();
    }
}

void MainWindow::onActiondeleteTriggered(QTableWidgetItem *item) {
    QDate d = ui->_calendar->selectedDate();
    int row = item->row();
    ui->_table->removeRow(row);
    Memo[d].erase(Memo[d].begin()+row + 1);
    activities.erase(activities.begin()+row + 1);
    // 处理 Action 1 被触发的逻辑
    //qDebug() << "Action 1 triggered";
}
void MainWindow::onActionaddTriggered(){
    int rowcount=ui->_table->rowCount();
    ui->_table->insertRow(rowcount);
    QDialog *dialogadd=new QDialog(this);
    dialogadd->setWindowTitle("Event Information");
    QVBoxLayout* layout = new QVBoxLayout(dialogadd);
    layout->addWidget(new QLabel("Start Time:"));
    QTimeEdit* stimeEdit = new QTimeEdit;
    stimeEdit->setDisplayFormat("hh:mm");
    layout->addWidget(stimeEdit);
    layout->addWidget(new QLabel("End Time:"));
    QTimeEdit* etimeEdit = new QTimeEdit;
    etimeEdit->setDisplayFormat("hh:mm");
    layout->addWidget(etimeEdit);
    layout->addWidget(new QLabel("Location:"));
    QComboBox* locationEdit = new QComboBox;
    for(auto nm:files.nameTint.keys()){
        locationEdit->addItem(nm);
    }

//    QLineEdit* locationEdit = new QLineEdit;
    layout->addWidget(locationEdit);
    layout->addWidget(new QLabel("Event:"));
    QLineEdit* eventEdit = new QLineEdit;
    layout->addWidget(eventEdit);
    QPushButton* saveButton = new QPushButton("Save");
    layout->addWidget(saveButton);
    connect(saveButton, &QPushButton::clicked, [=](){
            QTime stime = stimeEdit->time();
            QTime etime = etimeEdit->time();
            QString location = locationEdit->currentText();
            QString event = eventEdit->text();

            Event h;
            h.Sname=event;
            h.Sposition=location;
            h.begin=stime;
            h.end=etime;
            h.iposition=files.nameTint[location];
            if(Whetherclash(h.begin,h.end)){
                  InsertEvent(h);
            }


            dialogadd->close();
        });
    dialogadd->exec();
}
void MainWindow::onActioncancelTriggered(QTableWidgetItem *item){
    int row=item->row();
    activities[row].activate=0;
    for (int i = 0; i < ui->_table->columnCount(); i++) {
        QTableWidgetItem* item = ui->_table->item(row, i);
        if (item) {
            item->setBackground(QBrush(Qt::red));
        }
    }
}

bool MainWindow::Whetherclash(QTime begin,QTime end){
    for(auto it=activities.begin();it!=activities.end();it++){
        if((it->end>begin and it->begin<begin) or (it->begin<end and it->end>end)){
            QMessageBox msgBox(QMessageBox::Warning, "事件时间冲突", "该事件时间发生冲突,暂时不能给你明确的答复,请你自己衡量。", QMessageBox::NoButton, this);
             msgBox.setIcon(QMessageBox::Warning);
            QPushButton* continueButton = msgBox.addButton("继续添加", QMessageBox::AcceptRole);
            QPushButton* discardButton = msgBox.addButton("放弃", QMessageBox::RejectRole);
            msgBox.exec();

                if (msgBox.clickedButton() == continueButton) {
                    return 1; // 继续添加
                } else if (msgBox.clickedButton() == discardButton) {
                    return 0; // 放弃
                }

                return -1; // 未点击任何按钮
            break;
        }
    }
    return 1;
}

void MainWindow::Personalize(){
    configs->show();
}

void MainWindow::Submit(){
//    QDate date = QDate();
//    date = ui->_calendar->selectedDate();
//    int d = date.day();
//    int weekday = date.dayOfWeek();
    _pkumap = new PKUMap(this,activities,files.intTpos,files.intTpos[files.nameTint[configs->Origin()]]);
    ShowMap();
}

void MainWindow::AddRow(int row){
    ui->_table->insertRow(row);
}
void MainWindow::AddColumn(int column){
    ui->_table->insertColumn(column);
}
template<typename T>
void MainWindow::AddItem(int row,int column,T item){
    QTableWidgetItem *it = new QTableWidgetItem(item);
    ui->_table->setItem(row, column, it);
}

void MainWindow::AddEvent(Event& event){

    int nrow = ui->_table->rowCount();
    AddRow(nrow);
    AddItem(nrow,0,event.begin.toString());
    AddItem(nrow,1,event.end.toString());
    AddItem(nrow,2,event.Sname);
    AddItem(nrow,3,event.Sposition);
}

// 以下是直接操作储存事件的vector的接口
bool MainWindow::Available(QTime t){
    for(auto event:activities){
        if(event.begin <= t && event.end >= t){
            qDebug() << "not available";
            return 0;
        }
    }
    qDebug() << "available";
    return 1;
}

int MainWindow::FirstBefore(QTime t){
//    bool find = 0;
    int idx = -1;
    for(int i = activities.size() - 1; i >= 0; -- i){
        if(activities[i].end < t){
            idx = i;
            break;
        }
    }

    return idx;
}

void MainWindow::InsertEvent(Event& event){
    activities.push_back(event);
    SortEvent();
    files.saveUserInfo(event.iposition);
}

// 提供两个版本的删除课程接口
void MainWindow::DeleteEvent(Event& event){

}

void MainWindow::DeleteEvent(const QString& Sname){

}

// 不论是删除还是插入课程后都需要重新排序,如果我们只使用InsertEvent这个接口保证有序性，那这个函数不用实现
void MainWindow::SortEvent(){
    qDebug()<<activities.size()<<' '<<"kk"<<endl;
    std::sort(activities.begin(),activities.end());
    tableclear();
    for(Event u:activities){
        AddEvent(u);
    }
}

void MainWindow::GetSingle(QString name,QTime begin,QTime end,int type){
    QPair<int,int> curpos;
    QPair<int,int> nextpos;

    int idx = FirstBefore(begin);
    if(idx == -1){
        curpos = files.intTpos[files.nameTint[configs->Origin()]];
        if(activities.empty()){
            nextpos = files.intTpos[files.nameTint[configs->Origin()]];
        }
        else
        nextpos = files.intTpos[activities[0].iposition];
    }
    else if(idx != activities.size() - 1){
        curpos = files.intTpos[activities[idx].iposition];
        nextpos = files.intTpos[activities[idx + 1].iposition];
    }
    else{
        curpos = files.intTpos[activities[idx].iposition];
        nextpos = files.intTpos[files.nameTint[configs->Origin()]];
    }
    int iposition = 0;
    if(configs->Mode() == 0 && configs->Store()){
        iposition = Selection::PreferenceSelection(type);
    }
    else if(configs->Mode() == 0 && configs->Store() == 0){
        iposition = Selection::RandomSelection(type);
    }
    else{
        iposition = Selection::NearestSelection(type,curpos,nextpos);
    }
    QString Sposition = files.intTname[iposition];
    Event e(name,Sposition,iposition,begin,end);
    InsertEvent(e);

}

void MainWindow::GetStudy(){
    qDebug() << "Now add Study";

    int Study = Selection::RandomInt(0,7);
    if(configs->Mode() == 1){
        Study = 7;
    }

    if(Study & 1){
        QTime time = QTime::fromString("09:00:00","hh:mm:ss");
        QTime from = time;
        QTime limit = QTime::fromString("11:35:00","hh:mm:ss");
        while(!Available(from) && from <= limit){
            from = from.addSecs(1800);
        }
        QTime to = from;
        while(Available(to) && to <= limit){
            to = to.addSecs(1800);
        }
        to = to.addSecs(-1800);
        if(from < to){
            GetSingle("自习",from,to,1);
        }
    }

    if(Study & 2){
        QTime time = QTime::fromString("12:30:00","hh:mm:ss");
        QTime j = QTime::fromString("11:35:00","hh:mm::ss");
        if(Available(j)){
            time = time.addSecs(-1800);
        }
        QTime from = time;
        QTime limit = QTime::fromString("17:35:00","hh:mm:ss");
        while(!Available(from) && from <= limit){
            from = from.addSecs(1800);
        }
        QTime to = from;
        while(Available(to) && to <= limit){
            to = to.addSecs(1800);
        }
        to = to.addSecs(-1800);
        if(from < to){
            GetSingle("自习",from,to,1);
        }
    }

    if(Study & 4){
        QTime time = QTime::fromString("18:30:00","hh:mm:ss");
        QTime j = QTime::fromString("17:35:00","hh:mm::ss");
        if(Available(j)){
            time = time.addSecs(-1800);
        }
        QTime from = time;
        QTime limit = QTime::fromString("21:35:00","hh:mm:ss");
        if(configs->Mode() == 1){
            limit = limit.addSecs(3600);
        }
        while(!Available(from) && from <= limit){
            from = from.addSecs(1800);
        }
        QTime to = from;
        while(Available(to) && to <= limit){
           to = to.addSecs(1800);
        }
        to = to.addSecs(-1800);
        if(from < to){
            GetSingle("自习",from,to,1);
        }
    }
}

void MainWindow::GetFood(){
    qDebug() << "Now add Food";
    // 处理早餐
    if(configs->Breakfast()){
        // 考虑错时优化，例如有早八可以是7：30，没有早八可以是8：30
        QString timeStr = "08:00:01";
        QTime time = QTime::fromString(timeStr, "hh:mm:ss");
        if(Available(time)){
            timeStr = "08:30:00";
        }
        else{
            timeStr = "07:30:00";
        }
        QTime beg = QTime::fromString(timeStr,"hh:mm:ss");
        QTime ed = beg.addSecs(1200);

        GetSingle("早餐",beg,ed,2);

    }
    // 处理中餐
    if(configs->Lunch()){
        // 考虑错时优化，例如上午最后没有课可以是11：30，有课可以是12：00
        QString timeStr = "11:59:59";
        QTime time = QTime::fromString(timeStr, "hh:mm:ss");
        if(Available(time)){
            timeStr = "11:30:00";
        }
        else{
            timeStr = "12:00:00";
        }
        QTime beg = QTime::fromString(timeStr,"hh:mm:ss");
        QTime ed = beg.addSecs(1200);
        GetSingle("中餐",beg,ed,2);
    }



    // 处理晚餐
    if(configs->Dinner()){
        // 考虑错时优化，例如下午最后没有课可以是17：30，有课可以是17：00 / 18：00
        QString timeStr = "17:59:59";
        QTime time = QTime::fromString(timeStr, "hh:mm:ss");
        if(Available(time)){
            timeStr = "17:30:00";
        }
        else{
            timeStr = "18:00:00";
        }
        QTime beg = QTime::fromString(timeStr,"hh:mm:ss");
        QTime ed = beg.addSecs(1200);
        GetSingle("晚餐",beg,ed,2);

    }


}

void MainWindow::Bailan(){
    tableclear();
    int nrow = ui->_table->rowCount();
    AddRow(nrow);

    AddItem(nrow,0,"00:00:00");
    AddItem(nrow,1,"24:00:00");
    AddItem(nrow,2,"摆烂");
    AddItem(nrow,3,configs->Origin());
}

void MainWindow::GetActivity(){

}

// 这是是一个获取private变量的接口，之后整理代码的时候注意！
FileIO MainWindow::GetFile(){
    return files;
}
