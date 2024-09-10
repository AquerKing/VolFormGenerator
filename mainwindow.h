#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QLabel>
#include <QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

public slots:
    void on_comboBox_Session1_currentIndexChanged(int index);

private slots:
    void on_comboBox_Center_currentIndexChanged(int index);

    void on_pushButton_clicked();
    void on_pushButton_ChooseImage_clicked();

    void on_pushButton_Generate_clicked();

private:
    Ui::MainWindow *ui;

    int checkedSessionIndex;
    QString fileName;

    bool insertTextAtBookmark(QAxObject* doc, QString bookmarkName, QString text);
};
#endif // MAINWINDOW_H
