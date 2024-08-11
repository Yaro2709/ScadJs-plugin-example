var WshShell = new ActiveXObject("Wscript.Shell");

function Plugin_Execute(engine)
{
    var editor = engine.GetEditor();

    //Количество добавляемых узлов
    var quantityNode = 2;

    var baseNodeNum = editor.NodeAdd(quantityNode);

    //Создание узла 1
    var node_1 = {x: 0, y: 0, z: 0}
    editor.NodeUpdate(baseNodeNum, node_1)

    //Создание узла 2
    var node_2 = {x: 6, y: 0, z: 0}
    editor.NodeUpdate(baseNodeNum, node_2)



    // Связи
    var Bound = {
        Mask: 63,
        ListNode: [baseNodeNum, baseNodeNum + 1]
    }

    var bReplace = true;

    editor.SetBound(Bound, bReplace);


    // Элементы
    var quantityElem = 1;

    var baseElemNum = editor.ElemAdd(quantityElem);

    //Номер первого узла
    var startNumberNode = baseNodeNum;
    //Номер последнего узла
    var endNumberNode = baseNodeNum + 1;

    var elem = {
        TypeElem: 5,                                        //5 тип конечного элемента
        ListNode: [startNumberNode, endNumberNode]
    }

    editor.ElemUpdate(baseElemNum, elem);


    //Жесткости
    var Rigid = {
        Text: 'Balka',                                      //Назание жесткости
        ListElem: [baseElemNum],                            //Какому элементу присваивают жесткость
        Description: "STZ RUSSIAN pu_typ97 6 TMP 1.2e-005"  //Код жесткости (швеллер-стальной)
    }

    editor.RigidAdd(Rigid);


    //Нагрузка
    //Одно загружение
    var QuantityLoading = 1;

    var NumLoadingBeam = editor.LoadingAdd(QuantityLoading);

    var ForceElemBeam = {
        Qw: 16,                                             //Раномерная распределенная, общая система координат
        Qn: 3,                                              //Направлена по оси OZ
        ListData: [2000],                                   //Сила воздействия 2кН
        ListElem: [baseElemNum]                             //Задаем на наш элемент
    };
    editor.LoadingForceElemAdd(NumLoadingBeam, ForceElemBeam);

    var nameLoad = 'Test';                                  //Название загржуения
    var type = 1;                                           //Тип загружения
    var mode = 2;
    var longTime = 0.35;                                    //Доля длительности загружения
    var reliabilityFactor = 1.1;                            //Коэф. по надежности
    var Description =
        "Name=" +
        nameLoad +
        " Type=" +
        type +
        " Model=" +
        mode +
        " LongTime=" +
        longTime +
        " ReliabilityFactor" +
        reliabilityFactor;

    editor.LoadingSetDescription(NumLoadingBeam, Description);

    WshShell.Popup("Successfully!");
}
