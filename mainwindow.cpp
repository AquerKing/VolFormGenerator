#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QAxObject>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    // connect(ui->comboBox_Session1, &QComboBox::currentIndexChanged, this, &MainWindow::on_comboBox_Session1_currentIndexChanged);
    ui->comboBox_Center->addItem("请选择中心");
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_comboBox_Session1_currentIndexChanged(int index)
{
}


void MainWindow::on_comboBox_Center_currentIndexChanged(int index)
{
    ui->comboBox_Session1->clear();
    ui->comboBox_Session2->clear();

    ui->comboBox_Session1->addItem("请选择第一部门");
    ui->comboBox_Session2->addItem("请选择第二部门");

    switch(index)
    {
    case 1:
        ui->comboBox_Session1->addItem("办公室");
        ui->comboBox_Session1->addItem("宣传部");

        ui->comboBox_Session2->addItem("办公室");
        ui->comboBox_Session2->addItem("宣传部");
        break;
    case 2:
        ui->comboBox_Session1->addItem("文艺部");
        ui->comboBox_Session1->addItem("体育部");

        ui->comboBox_Session2->addItem("文艺部");
        ui->comboBox_Session2->addItem("体育部");
        break;
    case 3:
        ui->comboBox_Session1->addItem("权益部");
        ui->comboBox_Session1->addItem("生活部");

        ui->comboBox_Session2->addItem("权益部");
        ui->comboBox_Session2->addItem("生活部");
        break;
    case 4:
        ui->comboBox_Session1->addItem("学习部");
        ui->comboBox_Session1->addItem("自律委员会");

        ui->comboBox_Session2->addItem("学习部");
        ui->comboBox_Session2->addItem("自律委员会");
        break;
    case 5:
        ui->comboBox_Session1->addItem("综合组织部");
        ui->comboBox_Session1->addItem("青年发展部");
        ui->comboBox_Session1->addItem("班团建设部");

        ui->comboBox_Session2->addItem("综合组织部");
        ui->comboBox_Session2->addItem("青年发展部");
        ui->comboBox_Session2->addItem("班团建设部");
        break;
    case 6:
        ui->comboBox_Session1->addItem("实践服务部");
        ui->comboBox_Session1->addItem("组织活动部");

        ui->comboBox_Session2->addItem("实践服务部");
        ui->comboBox_Session2->addItem("组织活动部");
        break;
    case 7:
        ui->comboBox_Session1->addItem("媒体运营部");
        ui->comboBox_Session1->addItem("数媒编辑部");
        ui->comboBox_Session1->addItem("产品开发部");

        ui->comboBox_Session2->addItem("媒体运营部");
        ui->comboBox_Session2->addItem("数媒编辑部");
        ui->comboBox_Session2->addItem("产品开发部");
        break;
    case 8:
        ui->comboBox_Session1->addItem("竞赛服务部");
        ui->comboBox_Session1->addItem("宣传策划部");
        ui->comboBox_Session1->addItem("智能基座（华为）社团");
        ui->comboBox_Session1->addItem("CCF合肥工业大学学生分会执行委员部");

        ui->comboBox_Session2->addItem("竞赛服务部");
        ui->comboBox_Session2->addItem("宣传策划部");
        ui->comboBox_Session2->addItem("智能基座（华为）社团");
        ui->comboBox_Session2->addItem("CCF合肥工业大学学生分会执行委员部");
        break;
    case 9:
        ui->comboBox_Session1->addItem("培训部");
        ui->comboBox_Session1->addItem("活动部");

        ui->comboBox_Session2->addItem("培训部");
        ui->comboBox_Session2->addItem("活动部");
        break;
    }
}


void MainWindow::on_pushButton_clicked()
{
    QAxObject *word = new QAxObject("Word.Application");
    word->setProperty("Visible", false);

    QAxObject *documents = word->querySubObject("Documents");
    QAxObject *templateDoc = documents->querySubObject("Open(const QString&)", "C:/Users/aquer/Desktop/template.docx");

    // 替换书签
    QString bookmarks[17] = {"Name", "Gender", "ID", "Class", "DomNum", "Hometown",
                           "Role", "PhoneNum", "Center", "Session1", "Session2",
                           "AllAccept", "History", "SelfJudgement", "Realization",
                           "VolNum", "Photo"};

    bool firstVol = false;
    bool secondVol = false;
    for (int i = 0; i < 16; i++)
    {
        QString text;
        QAxObject *bookmark = templateDoc->querySubObject("Bookmarks(const QString&)", bookmarks[i]);
        if (!bookmark->isNull()) {
            bookmark->dynamicCall("Select()");
            QAxObject *range = bookmark->querySubObject("Range()");

            switch(i)
            {
            case 0:
                text = this->ui->lineEdit_Name->text();
                break;
            case 1:
                text = this->ui->comboBox_Gender->currentText();
                break;
            case 2:
                text = this->ui->lineEdit_ID->text();
                break;
            case 3:
                text = this->ui->lineEdit_Class->text();
                break;
            case 4:
                text = this->ui->lineEdit_DomNum->text();
                break;
            case 5:
                text = this->ui->lineEdit_Hometown->text();
                break;
            case 6:
                text = this->ui->comboBox_Role->currentText();
                break;
            case 7:
                text = this->ui->lineEdit_PhoneNum->text();
                break;
            case 8:
                text = this->ui->comboBox_Center->currentText();
                break;
            case 9:
                if (this->ui->comboBox_Session1->currentIndex() == 0) {
                    return;
                }
                text = this->ui->comboBox_Session1->currentText();
                firstVol = true;
                break;
            case 10:
                if (this->ui->comboBox_Session2->currentIndex() == 0 &&
                    this->ui->comboBox_Session1->currentIndex() == 0) {
                    return;
                }
                text = this->ui->comboBox_Session2->currentText();
                secondVol = true;
                break;
            case 11:
                text = this->ui->checkBox_AllAccept->isChecked() ? "是" : "否";
                break;
            case 12:
                text = this->ui->plainTextEdit_History->toPlainText();
                break;
            case 13:
                text = this->ui->plainTextEdit_SelfJudgement->toPlainText();
                break;
            case 14:
                text = this->ui->plainTextEdit_Realization->toPlainText();
                break;
            case 15:
                text = this->ui->comboBox_VolNum->currentIndex() == 0 ? "一" : "二";
                break;
            }

            range->setProperty("Text", text);
        }
    }

    QString newDocName = QString("第%1志愿-%2-%3%4-%5-%6").arg(this->ui->comboBox_VolNum->currentIndex() == 0 ? "一" : "二")
                             .arg(this->ui->comboBox_Center->currentText()).arg(secondVol ? QString("%1-%2").arg(this->ui->comboBox_Session1->currentText(), this->ui->comboBox_Session2->currentText()) : QString("%1-").arg(this->ui->comboBox_Session1->currentText()))
        .arg(this->ui->lineEdit_ID->text()).arg(this->ui->lineEdit_Name->text()).arg(this->ui->lineEdit_Class->text());
    templateDoc->dynamicCall("SaveAs(const QString&)", QString("./%1.docx").arg(newDocName));

    templateDoc->dynamicCall("Close()");
    word->dynamicCall("Quit()");

    delete word;
}

