var WshShell = new ActiveXObject("Wscript.Shell");

function Plugin_Execute(engine)
{
    //Получение усилий в элемента
    var result = engine.GetResult();

    //Номер правой части (нелинейные расчеты, для статики 1)
    var NumRHS = 1;
    var NumElem = 1;
    //Номер воздействия
    var NumAction =1;
    //Номер шага (нелинейные расчеты, для статики 1)
    var NumFixedStep = 1;

    var Efforts = {
        //Количество элементов в массиве ListUs
        QuantityUs:null,
        //массив типов напряжений/усилий вычисленных для элемента с номером NumElem (дополнение 6)
        ListUs:null,
        QuantityData:null,
        //Для каждой точки выдачи напряжений/усилий - значки в порядкею
        ListData:null
    }

    //напряжение/усилия
    result.GetEfforts(NumAction, NumFixedStep, NumRHS, NumElem, Efforts);

    Efforts.ListUs = Efforts.ListUs.toArray();
    Efforts.ListData = Efforts.ListData.toArray();

    var effortArr = [];
    effortArr.push(Efforts.ListData.slice(0, Efforts.QuantityUs));
    effortArr.push(Efforts.ListData.slice(Efforts.QuantityUs, Efforts.QuantityUs + Efforts.QuantityUs));
    effortArr.push(Efforts.ListData.slice(Efforts.QuantityUs + Efforts.QuantityUs, Efforts.QuantityData));

    //Excel - отчет
    var excel = new ActiveXObject("Excel.Application");
    excel.Visible = true;
    excel.Workbooks.Add;
    var worksheet = excel.Workbooks.Application.Sheets.Item(1);

    var range = worksheet.Range("A1:A2");
    range.Merge();
    excel.Cells(1,1).Value = "Seth";

    var startRow = 3;
    var startColumn = 2;

    for(var i = 0; i < 3; i++){
        excel.Cells(startRow + i, 1).Value = i + 1;

        for(var j = 0; j < effortArr[i].length; j++){
            excel.Cells(startRow + i, startColumn + j).Value = effortArr[i][j];
        }
    }

    var effortNameArr = ["N", "Mk", "My", "Qz", "Mz", "Qy", "Rx", "Ry", "Rz"];

    for(var i = 0; i < effortNameArr.length; i++){
        excel.Cells(startRow - 1, startColumn + i).Value = effortNameArr[i];
    }

    var range_2 = worksheet.Range("B1:J1");
    range_2.Merge();
    excel.Cells(1,2).Value = "Element" + " forces in " + NumElem;

    WshShell.Popup("Successfully!");
}
