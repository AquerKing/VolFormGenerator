#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QAxObject>
#include <QFileDialog>

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

}

void MainWindow::on_pushButton_ChooseImage_clicked()
{
    this->fileName = QFileDialog::getOpenFileName(this, tr("选择照片"), QDir::homePath(),
                                                    tr("Image File (*.png | *.jpg | *.jpeg)"));
    if (!fileName.isEmpty())
    {
        QPixmap pixmap(fileName);
        this->ui->label_Image->setPixmap(pixmap.scaled(this->ui->label_Image->size(), Qt::KeepAspectRatio, Qt::SmoothTransformation));
    }
}


void MainWindow::on_pushButton_Generate_clicked()
{
    ui->label->setText("正在生成中，请稍等...期间窗口未响应属正常现象");
    ui->label->repaint();
    QAxObject *word = new QAxObject("Word.Application");
    word->setProperty("Visible", false);

    QAxObject *documents = word->querySubObject("Documents");
    QAxObject *templateDoc = documents->querySubObject("Open(const QString&)", "C:/Users/aquer/Desktop/template.docx");

    // 替换书签
    QString bookmarks[16] = {"Name", "Gender", "ID", "Class", "DomNum", "Hometown",
                             "Role", "PhoneNum", "Center", "Session1", "Session2",
                             "AllAccept", "History", "SelfJudgement", "Realization",
                             "VolNum"};

    // bool firstVol = false;
    bool secondVol = false;

    this->insertTextAtBookmark(templateDoc, "Name", this->ui->lineEdit_Name->text());
    this->insertTextAtBookmark(templateDoc, "Gender", this->ui->comboBox_Gender->currentText());
    this->insertTextAtBookmark(templateDoc, "ID", this->ui->lineEdit_ID->text());
    this->insertTextAtBookmark(templateDoc, "Class", this->ui->lineEdit_Class->text());
    this->insertTextAtBookmark(templateDoc, "DomNum", this->ui->lineEdit_DomNum->text());
    this->insertTextAtBookmark(templateDoc, "Hometown", this->ui->lineEdit_Hometown->text());
    this->insertTextAtBookmark(templateDoc, "Role", this->ui->comboBox_Role->currentText());
    this->insertTextAtBookmark(templateDoc, "PhoneNum", this->ui->lineEdit_PhoneNum->text());
    this->insertTextAtBookmark(templateDoc, "Center", this->ui->comboBox_Center->currentText());

    this->insertTextAtBookmark(templateDoc, "Session1", this->ui->comboBox_Session1->currentText());
    if (this->ui->comboBox_Session2->currentIndex() != 0 &&
        this->ui->comboBox_Session2->currentIndex() != this->ui->comboBox_Session1->currentIndex())
    {
        this->insertTextAtBookmark(templateDoc, "Session2", this->ui->comboBox_Session2->currentText());
        secondVol = true;
    }
    this->insertTextAtBookmark(templateDoc, "AllAccept", this->ui->checkBox_AllAccept->isChecked() ? "是" : "否");
    this->insertTextAtBookmark(templateDoc, "History", this->ui->plainTextEdit_History->toPlainText());
    this->insertTextAtBookmark(templateDoc, "SelfJudgement", this->ui->plainTextEdit_SelfJudgement->toPlainText());
    this->insertTextAtBookmark(templateDoc, "Realization", this->ui->plainTextEdit_Realization->toPlainText());
    this->insertTextAtBookmark(templateDoc, "VolNum", this->ui->comboBox_VolNum->currentIndex() == 0 ? "一" : "二");


    qDebug() << "File Path: " << fileName;
    // 插入图片
    if (fileName.isEmpty())
    {
        ui->label->setText("生成失败：未选择照片");
        delete word;
        return;
    }
    QAxObject *bookmark = templateDoc->querySubObject("Bookmarks(const QString&)", "Photo");
    if (bookmark) {
        bookmark->dynamicCall("Select()");
        QAxObject *inlineShapes = templateDoc->querySubObject("InlineShapes");
        QVariantList args;
        args << fileName << false << true;
        inlineShapes->dynamicCall("AddPicture(const QString&, QVariant, QVariant)", args);
        QAxObject *shape = inlineShapes->querySubObject("Item(int)", 1);
        shape->setProperty("LockAspectRatio", true);
    } else {
        qDebug() << "Bookmark not found";
    }

    QString newDocName = QString("第%1志愿-%2-%3%4-%5-%6").arg(this->ui->comboBox_VolNum->currentIndex() == 0 ? "一" : "二")
                             .arg(this->ui->comboBox_Center->currentText()).arg(secondVol ? QString("%1-%2-").arg(this->ui->comboBox_Session1->currentText(), this->ui->comboBox_Session2->currentText()) : QString("%1-").arg(this->ui->comboBox_Session1->currentText()))
                             .arg(this->ui->lineEdit_ID->text()).arg(this->ui->lineEdit_Name->text()).arg(this->ui->lineEdit_Class->text());
    templateDoc->dynamicCall("SaveAs(const QString&)", QString("C:/Users/aquer/Desktop/%1.docx").arg(newDocName));

    templateDoc->dynamicCall("Close()");
    word->dynamicCall("Quit()");

    ui->label->setText("生成成功");
    delete word;
}

bool MainWindow::insertTextAtBookmark(QAxObject* doc, QString bookmarkName, QString text)
{
    if (doc == nullptr)
        return false;
    QAxObject *bookmark = doc->querySubObject("Bookmarks(const QString&)", bookmarkName);
    if (!bookmark->isNull()) {
        bookmark->dynamicCall("Select()");
        QAxObject *range = bookmark->querySubObject("Range()");
        range->setProperty("Text", text);
    }
}

