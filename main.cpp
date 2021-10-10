#include <QCoreApplication>
#include <QFile>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QJsonDocument>
#include <QDebug>
#include <QString>
#include <QVariant>
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

QString readDocCell (int row, int coll, QXlsx::Document *xlsx)
{
    QString cell;

    if (xlsx->load())
    {
        QXlsx::Cell *docCell = xlsx->cellAt(row, coll);

        if (docCell != NULL)
        {
            cell = docCell->readValue().toString();
        }
    }

return cell;
}

QJsonObject createJsonObject (QXlsx::Document *xlsx)
{
    QJsonObject jsonObj;
    QJsonValue jsonValue;
    QString jsonKey;

    int counter = 1;

    do
    {
        jsonValue = readDocCell(counter, 2, xlsx);
        jsonKey = readDocCell(counter, 1, xlsx);

        QJsonObject::Iterator jsonIter;
        bool skipFlag = false;

        if (jsonObj.find(jsonKey) != jsonObj.end())
        {
                QJsonArray jsonArray;
                QString jsonArrayKey;

                jsonArrayKey = jsonKey;
                jsonArray.append(jsonValue);

                while (jsonObj.find(jsonKey) != jsonObj.end())
                {
                    jsonArray.append(jsonObj.find(jsonKey).value());
                    jsonObj.erase(jsonObj.find(jsonKey));
                }
                jsonObj.insert(jsonArrayKey, jsonArray);
                counter++;
                continue;
        }


        jsonObj.insert(jsonKey, jsonValue);
        counter++;
    }
    while (jsonValue.isNull() != true && jsonKey != NULL);

    jsonObj.erase(jsonObj.begin());

    return jsonObj;
}


int main(int argc, char *argv[])
{
//    QCoreApplication a(argc, argv);

//    if (argc > 3)
//    {
//        qDebug() << "Error: too much arguments to call" << Qt::endl;
//        return 0;
//    }

    QFile file(argv[1]); // "/home/u34/Рабочий стол/QXlsx_proj/Output.json"
    QXlsx::Document xlsx("/home/u34/Рабочий стол/QXlsx_proj/Test.xlsx");
    QJsonDocument jsonDoc;

        jsonDoc.setObject(createJsonObject(&xlsx));
        file.open(QIODevice::WriteOnly);
        file.write(jsonDoc.toJson(QJsonDocument::Indented));
        file.close();

    return 0;
    // return a.exec();
}
