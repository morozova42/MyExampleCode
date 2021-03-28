using Giprosintez.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace AddPodpisDWG
{
    [Export(typeof(IGroupAction)),
        ExportMetadata("ActionName", "Вставка блоков подписей в чертежи vvvvvvvv"),
        ExportMetadata("ActionID", "05"),
        ExportMetadata("ActionDescription", "Вставка подписей в чертежи")]
    public class AddPodpisDWGClass : IGroupAction
    {
        #region Данные модуля
        Autodesk.AutoCAD.Interop.AcadApplication acad;
        Autodesk.AutoCAD.Interop.AcadDocument curDraw;
        Dictionary<string, string> Blocks;
        Dictionary<string, string> BlocksLeft;
        Dictionary<string, List<string>> BlocksDublir;
        /// <summary>
        /// Словарь фамилий на чертеже с точками вставки подписей
        /// </summary>
        SortedDictionary<Point, string> Positions;
        /// <summary>
        /// Словарь общих для всех чертежей дублирующихся фамилий
        /// </summary>
        Dictionary<string, string> UsedDublFIOs;
        SortedDictionary<int, Izm> CommonIzmDict;
        /// <summary>
        /// Словарь № изменений и координат подписей к ним
        /// </summary>
        SortedDictionary<int, Izm> CurrentIzmDict;
        public bool FlagAll;
        bool FlagLists;
        bool ZamOrNov;
        int CurrentIzmNumber;
        public string FIOIzm;
        public string FIODublikat;
        public string DateTimeRazrab;
        public string DateTimeIzm;
        private Dictionary<string, List<double>> StampNames;
        dynamic explodedObjects = null;
        IAcadBlockReference StampBlock = null;
        IAcadBlockReference LeftStampBlock = null;
        IAcadBlockReference StampDublicat = null;
        AcadSelectionSet MySelection;

        /// <summary>
        /// Список исходных файлов
        /// </summary>
        public List<string> Source { get; set; }

        /// <summary>
        /// Список результирующих файлов
        /// </summary>
        public List<string> Result { get; set; }

        /// <summary>
        /// Список входных параметров в виде пар "Параметр=значение"
        /// </summary>
        public List<string> InputParam { get; set; }

        /// <summary>
        /// Список выходных параметров в виде пар "Параметр=значение"
        /// </summary>
        public List<string> OutputParam { get; set; }

        /// <summary>
        /// Список сообщений действия
        /// </summary>
        public List<string> Message { get; set; }

        /// <summary>
        /// Свойство - флаг требования AutoCAD для работы
        /// </summary>
        public bool AcadRequired { get { return true; } }

        /// <summary>
        /// Свойство - ссылка на работающий Автокад
        /// </summary>
        public AcadApplication Acad { get { return acad; } set { acad = value; } }

        #endregion

        /// <summary>
        /// Конструктор класса
        /// </summary>
        public AddPodpisDWGClass()
        {
        }

        /// <summary>
        /// Заполнить словарь используемых штампов
        /// </summary>
        public void GetStampNames()
        {
            if (StampNames == null)
            {
                StampNames = new Dictionary<string, List<double>>
                {
                    { "_ST-2", new List<double> { 145, 35 } },
                    { "_ST-2P-new-izom", new List<double> { 145, 5 } },
                    { "_ST-2P-izom", new List<double> { 145, 5 } },
                    { "_ST-2dpt-Logotype", new List<double> { 145, 60 } },
                    { "_ST-3-Logotype", new List<double> { 145, 60 } },
                    { "_ST-3", new List<double> { 145, 60 } },
                    { "_ST-2-Logotype", new List<double> { 145, 50 } },
                    { "_ST-1-Logotype", new List<double> { 210, 5 } },
                    { "_ST-7-Logotype", new List<double> { 145, 35 } },
                    { "_ST-7", new List<double> { 145, 35 } },
                    { "_st-1-Logotype", new List<double> { 145, 35 } },
                    { "_st-4", new List<double> { 145, 35 } },
                    { "_ST-PKO-2P_", new List<double> { 145, 5 } },
                    { "_ST-SHTAMP-GRNO", new List<double> { 145, 35 } }
                };  //Список имён используемых штампов
            }
        }

        /// <summary>
        /// Заполнить словари блоков подписей
        /// </summary>
        public void GetBlocks()
        {
            UsedDublFIOs = new Dictionary<string, string>();
            Blocks = new Dictionary<string, string>();
            string Path = "S:\\ZOffice\\IT\\Presentation_&_Animation\\Подписи\\Блоки\\Чертежи";
            Blocks = GetBlocks(Path);

            Path = "S:\\ZOffice\\IT\\Presentation_&_Animation\\Подписи\\Блоки\\Боковой_штамп";

            BlocksLeft = new Dictionary<string, string>();
            BlocksLeft = GetBlocks(Path);

            BlocksDublir = new Dictionary<string, List<string>>();

            foreach (var blockItem in Blocks)  //Ищем среди блоков подписи с пробелами, заносим в отдельный словарь
            {
                if (blockItem.Key.Contains(" "))
                {
                    if (!BlocksDublir.ContainsKey(blockItem.Key.Split()[0]))
                    {
                        BlocksDublir.Add(blockItem.Key.Split()[0], new List<string> { blockItem.Key });
                    }
                    else
                    {
                        BlocksDublir[blockItem.Key.Split()[0]].Add(blockItem.Key);
                    }
                }
            }
        }

        /// <summary>
        /// Возвращает словарь блоков подписей
        /// </summary>
        /// <param name="path">Исходная папка</param>
        /// <returns></returns>
        private Dictionary<string, string> GetBlocks(string path)
        {
            Dictionary<string, string> Result = new Dictionary<string, string>();

            //списки файлов
            FileInfo[] fiAcadList;

            if (path.Length > 0)
            {
                SearchOption dirOption;//опция поиска
                DirectoryInfo dirInfo = new DirectoryInfo(path);//каталог поиска

                //определить параметр поиска файлов (с подкаталогами или без)
                dirOption = SearchOption.AllDirectories;

                fiAcadList = dirInfo.GetFiles("*.dwg", dirOption);
                foreach (FileInfo fi in fiAcadList)
                {
                    Result.Add(fi.Name.Substring(0, fi.Name.Length - 4), fi.FullName);
                }
                fiAcadList = dirInfo.GetFiles("*.dxf", dirOption);
                foreach (FileInfo fi in fiAcadList)
                {
                    Result.Add(fi.Name.Substring(0, fi.Name.Length - 4), fi.FullName);
                }
            }
            return Result;
        }

        /// <summary>
        /// Метод обработки списка файлов
        /// </summary>
        /// <param name="files"></param>
        /// <returns>Число обработанных файлов</returns>
        public int RunAction(List<string> files)
        {
            Message = new List<string>();
            try
            {
                return 2;
            }
            catch
            {
                return 0;
            }
        }

        public bool RunActionFile(string fileName)
        {
            Message = new List<string>();
            try
            {
                acad.WindowState = AcWindowState.acMax;
                // Проверка расширения файла
                if (!fileName.EndsWith(".dwg"))
                {
                    Message.Add("Неверный формат файла");
                    return false;
                }
                else
                {
                    curDraw = acad.ActiveDocument;
                    MySelection = curDraw.SelectionSets.Add("CurrentSet");
                    ZamOrNov = false;
                    bool success = AddBlocksPodpis();
                    if (!success) return false;
                }
                //curDraw.PurgeAll();

                curDraw.Regen(AcRegenType.acAllViewports);
                curDraw.Save();
                curDraw.SendCommand("(command \"_sdi\"  \"0\") ");
                return true;
            }
            catch (Exception ex)
            {
                Message.Add(ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Добавление подписей в чертёж с изменениями
        /// </summary>
        private bool AddBlocksPodpis()
        {
            curDraw.SendCommand("(command \"_zoom\" \"_e\") ");
            // Делаем активным слой Текст или Фамилии
            foreach (AcadLayer layer in curDraw.Layers)
            {
                if (layer.Name.ToUpper() == "ТЕКСТ" || layer.Name.ToUpper() == "ФАМИЛИИ")
                {
                    curDraw.ActiveLayer = layer;
                }
            }
            StampBlock = null;

            // Ищем штамп и записываем координаты точки, где у него вставляются подписи для измов
            double[] PtIzmPodpis = FindRightStampPodpisPoint();
            // Ищем блок форматной рамки и записываем точку бокового штампа
            double[] RamkaInsertionPoint = FindRamkaPoint(); //FindLeftStampPodpisPoint();

            DeletePodpisi();
            Positions = new SortedDictionary<Point, string>();
            // Собираем все фамилии на чертеже в словарь
            Positions = CollectAllFIO();
            if (Positions.Count == 0 && StampBlock.Name != "_ST-2P-new-izom" && StampBlock.Name != "_ST-2P-izom" && StampBlock.Name != "_ST-PKO-2P_")
            {
                Message.Add("В чертеже " + curDraw.Name + " фамилии не найдены");
            }
            else
            {
                //разделение словаря на левые (вертикальные) и правые подписи
                SortedDictionary<Point, string> PosLeft = new SortedDictionary<Point, string>();   
                SortedDictionary<Point, string> PosRigth = new SortedDictionary<Point, string>();

                // Для каждой точки вставки подписи проверяем координаты на соответствие штампам
                foreach (var item in Positions)
                {
                    double x = item.Key.X;
                    double y = item.Key.Y;
                    if ((LeftStampBlock != null) && ((x < RamkaInsertionPoint[0]) && (y > RamkaInsertionPoint[1])))
                    {
                        PosLeft.Add(item.Key, item.Value + "_бш");
                    }
                    else
                    {
                        PosRigth.Add(item.Key, item.Value);
                    }
                }

                bool success = Insert_Izm_Podpis_Right(PtIzmPodpis);
                if (!success)
                {
                    return false;
                }
                List<dynamic> tempEntities = Insert_Podpis_With_Familiya_Right(PosRigth);
                Delete_Temp(tempEntities);
                tempEntities = Insert_Podpis_With_Familiya_Left(PosLeft);
                Delete_Temp(tempEntities);
            }
            return true;
        }

        /// <summary>
        /// Удалить временные объекты
        /// </summary>
        /// <param name="tempEntities"></param>
        private void Delete_Temp(List<dynamic> tempEntities)
        {
            foreach (var item in tempEntities)
            {
                foreach (AcadEntity item1 in item)
                {
                    try
                    {
                        item1.Delete();
                    }
                    catch (Exception)
                    { }
                }
            }
        }

        /// <summary>
        /// Поиск блока штампа, штампа дубликат и определение точки для поиска подписи
        /// </summary>
        /// <returns>Точка х, у, z</returns>
        private double[] FindRightStampPodpisPoint()
        {
            double[] resultPoint = new double[] { 0, 0, 0 };
            short[] filterTypeBlock = { 0, 2 };
            object[] filterDataBlock = { "INSERT", "" };

            // Ищем на чертеже блоки по каждому имени в словаре штампов
            foreach (string stampName in StampNames.Keys)
            {
                filterDataBlock[1] = stampName;
                MySelection.Clear();
                MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeBlock, FilterData: filterDataBlock);
                if (MySelection.Count > 0)
                {
                    StampBlock = (IAcadBlockReference)MySelection.Item(0);
                    //Запись координат точки вставки подписи в измы
                    resultPoint = StampBlock.InsertionPoint;
                    resultPoint[0] = resultPoint[0] - StampNames[stampName][0];
                    resultPoint[1] = resultPoint[1] + StampNames[stampName][1];
                    break;
                }
            }
            // Если штамп не найден
            if (StampBlock == null)
            {
                Message.Add("В чертеже " + curDraw.Name + " не найден угловой штамп");
            }
            //Поиск штампа Дубликат
            filterDataBlock[1] = "dublikat";
            MySelection.Clear();
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeBlock, FilterData: filterDataBlock);
            if (MySelection.Count > 0)
            {
                StampDublicat = (IAcadBlockReference)MySelection.Item(0);
            }
            return resultPoint;
        }

        /// <summary>
        /// Поиск блока рамки и координаты точки края бокового штампа
        /// </summary>
        /// <returns></returns>
        private double[] FindRamkaPoint()
        {
            // Ищем все блоки на чертеже
            double[] resultPoint = new double[] { 0, 0, 0 };
            short[] filterTypeBlock = { 0, 2 };
            object[] filterDataBlock = { "INSERT" };
            MySelection.Clear();
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeBlock, FilterData: filterDataBlock);
            foreach (AcadEntity item in MySelection)
            {
                IAcadBlockReference block = (IAcadBlockReference)item;
                //Проверяем имя каждого найденного блока
                if ((block.Name.StartsWith("A") || block.Name.StartsWith("_MT_")) && !block.Name.Contains("$"))
                {
                    if (block.Name.Length < 4 || (block.Name.Contains("x") && block.Name.Length < 6) || block.Name.Contains("GRNO"))
                    {
                        LeftStampBlock = (IAcadBlockReference)item;
                        resultPoint = LeftStampBlock.InsertionPoint;
                        //Записываем координаты края левого бокового штампа
                        resultPoint[0] += 20;
                        resultPoint[1] += 90;
                    }
                }
            }
            return resultPoint;
        }

        /// <summary>
        /// Определение положения левого (вертикального) штампа
        /// </summary>
        /// <returns></returns>
        private double[] FindLeftStampPodpisPoint()
        {
            double[] resultPoint = new double[] { 0, 0, 0 };
            short[] filterTypeBlockLeft = { 0, 2 };
            object[] filterDataBlockLeft = { "INSERT", "A$C412B1E25" };

            MySelection.Clear();
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeBlockLeft, FilterData: filterDataBlockLeft);

            if (MySelection.Count > 0)
            {
                IAcadBlockReference LeftStampBlock = (IAcadBlockReference)MySelection.Item(0);
                resultPoint[0] = LeftStampBlock.InsertionPoint[0];
                resultPoint[1] = LeftStampBlock.InsertionPoint[1];
            }
            return resultPoint;
        }

        /// <summary>
        /// Найти и удалить имеющиеся блоки подписей
        /// </summary>
        private void DeletePodpisi()
        {
            //Выбираем все вхождения блоков
            short[] filterTypePodpis = { 0 };
            object[] filterDataPodpis = { "INSERT" };
            MySelection.Clear();
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypePodpis, FilterData: filterDataPodpis);

            //удаляем те, что встречаются в блоках фамилий
            foreach (IAcadBlockReference itemBlock in MySelection)  
            {
                if (Blocks.ContainsKey(itemBlock.Name))
                {
                    itemBlock.Delete();
                }
                else if (BlocksLeft.ContainsKey(itemBlock.Name))
                {
                    itemBlock.Delete();
                }
            }

            try
            {
                var depend = curDraw.FileDependencies;
                foreach (AcadFileDependency item in depend)
                {
                    string FileName = item.FullFileName;

                    if (FileName.Contains(".png") && FileName.Contains(@"ZOffice\IT\Presentation_&_Animation\"))
                    {
                        curDraw.FileDependencies.RemoveEntry(item.Index, true);
                        int count = curDraw.FileDependencies.Count;
                    }
                }
            }
            catch (Exception ex)
            {

            }

            foreach (var item in curDraw.Dictionaries)
            {
                try
                {
                    AcadDictionary AD = (AcadDictionary)item;
                    if (AD.Name == "ACAD_IMAGE_DICT")
                    {
                        foreach (AcadObject item1 in AD)
                        {
                            string PicName = AD.GetName(item1);
                            if (Blocks.ContainsKey(PicName) || BlocksLeft.ContainsKey(PicName))
                            {
                                if (curDraw.Name.Contains(PicName) == false)
                                {
                                    try
                                    {
                                        item1.Delete();
                                    }
                                    catch (System.Exception ex)
                                    {
                                        Message.Add(curDraw.Name + "- Ошибка: " + ex.Message);
                                    }
                                }
                            }
                        }
                        break;
                    }
                }
                catch (System.Exception)
                {

                }
            }
            curDraw.PurgeAll();
        }

        /// <summary>
        /// Поиск подписей и запись их координат в словарь
        /// </summary>
        /// <returns></returns>
        private SortedDictionary<Point, string> CollectAllFIO()
        {
            Positions = new SortedDictionary<Point, string>();

            //ищем весь текст на чертеже
            short[] filterTypeText = { 0 };
            object[] filterDataText = { "TEXT" };
            MySelection.Clear();
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeText, FilterData: filterDataText);
            filterDataText[0] = "MTEXT";
            MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeText, FilterData: filterDataText);           

            // Проверяем каждое вхождение текста
            foreach (AcadEntity item in MySelection)
            {
                IAcadMText MText1;
                IAcadText Text1;
                string Familiya = "Фамилия не задана";
                Point FIOPosition = new Point(0, 0, 0);
                
            //записываем текст и точку вставки текста
                if (item.ObjectName == "AcDbText")                                  
                {
                    Text1 = (IAcadText)item;
                    Familiya = Text1.TextString.Trim();
                    FIOPosition = new Point(Text1.InsertionPoint);
                }
                else if (item.ObjectName == "AcDbMText")                            
                {
                    MText1 = (IAcadMText)item;
                    if (MText1.TextString.Contains(';'))
                    {
                        Familiya = MText1.TextString.Trim().Split(';')[1].Split('}')[0];
                    }
                    else
                    {
                        Familiya = MText1.TextString.Trim();
                    }
                    FIOPosition = new Point(MText1.InsertionPoint);
                }

            //если текст - это фамилия из блоков, записываем её с точкой вставки в словарь
                if (Blocks.ContainsKey(Familiya))                                       
                {
                    if (!Positions.ContainsKey(FIOPosition))
                    {
                        Positions.Add(FIOPosition, Familiya);
                    }
                }
            //если фамилии нет в блоках, но есть среди повторяющихся фамилий
                else if (BlocksDublir.ContainsKey(Familiya))                       
                {
                // если в использованных дубликатах нет, спрашиваем у пользователя фамилию
                    if (!UsedDublFIOs.ContainsKey(Familiya))                       
                    {
                        RequestFIO(Familiya, FIOPosition);
                    }
                // если такая уже была, добавляем из использованных
                    else if (UsedDublFIOs.ContainsKey(Familiya))                   
                    {
                        Positions.Add(FIOPosition, UsedDublFIOs[Familiya]);
                    }
                }
            }
            return Positions;
        }

        private void RequestFIO(string Familiya, Point FIOPosition)
        {
            // Оформляем окно
            DublFamiliya dublForm = new DublFamiliya();
            dublForm.label1.Text = "На чертеже " + curDraw.Name + " обнаружена фамилия '" + Familiya + "'.";
            dublForm.radioButton1.Text = BlocksDublir[Familiya][0];
            dublForm.radioButton2.Text = BlocksDublir[Familiya][1];

            if (dublForm.ShowDialog() == DialogResult.OK)
            {
                string Temp = dublForm.SelectedFIO;

                // Если эта точка вставки подписи ещё не обработана
                if ( !Positions.ContainsKey(FIOPosition))
                {
                    // Заносим в словарь
                    Positions.Add(FIOPosition, Temp);
                    // Если выбрано использование этой фамилии во всех чертежах
                    if (dublForm.FlagAll)
                    {
                        // Добавляем в словарь общих дублирующихся фамилий
                        UsedDublFIOs.Add(Familiya, Temp);
                    }
                }
            }
        }

        /// <summary>
        /// Вычисление координат ближайших линий и вставка блока для правого штампа
        /// </summary>
        /// <param name="PosRigth">Список координат и фамилий чертежа</param>
        private List<dynamic> Insert_Podpis_With_Familiya_Right(SortedDictionary<Point, string> PosRigth)
        {
            List<dynamic> ListExplodedObjects = new List<dynamic>();
            List<IAcadLine> ListLine = new List<IAcadLine>();                   //список отрезков
            if (StampBlock != null)
            {
                explodedObjects = StampBlock.Explode();
                curDraw.Regen(AcRegenType.acAllViewports);
            }
            if (explodedObjects != null)                            //Если блок штампа был расчленён
            {
                foreach (AcadEntity item in explodedObjects)
                {
                    if (item.ObjectName == "AcDbLine")              //отрезок добавляем в список
                    {
                        ListLine.Add((IAcadLine)item);
                    }
                    else if (item.ObjectName == "AcDbPolyline")              //полилинию рачленяем
                    {
                        AcadLWPolyline LWPolyline = (AcadLWPolyline)item;

                        dynamic explodedObjects2 = LWPolyline.Explode();
                        foreach (AcadEntity item1 in explodedObjects2)
                        {
                            if (item1.ObjectName == "AcDbLine")             //получившиеся отрезки прибавляем к списку
                            {
                                ListLine.Add((IAcadLine)item1);
                            }
                        }
                        ListExplodedObjects.Add(explodedObjects2);
                    }
                }
            }
            else  //если блок не расчленён
            {
                short[] filterTypeLine = { 0 };
                object[] filterDataLine = { "LWPOLYLINE" };
                MySelection.Clear();
                MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeLine, FilterData: filterDataLine);

                foreach (IAcadLWPolyline item in MySelection)
                {
                    dynamic explodedObjects2 = null;
                    explodedObjects2 = item.Explode();
                    ListExplodedObjects.Add(explodedObjects2);
                }
                curDraw.Regen(AcRegenType.acAllViewports);

                filterDataLine[0] = "LINE";
                MySelection.Clear();
                MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeLine, FilterData: filterDataLine);
                foreach (IAcadLine item in MySelection)
                {
                    ListLine.Add(item);
                }
            }
            curDraw.Regen(AcRegenType.acAllViewports);

            double[] Point1 = new double[] { 0, 0, 0 };   //границы прямоугольника, в который должно попасть пересечение отрезков штампа (будущая т.вставки блока подписи)
            double[] Point2 = new double[] { 0, 0, 0 };
            double MaxY = -100;

            foreach (var itemPosition in PosRigth)
            {
                double Xadd = 0;                    //точка вставки подписи
                double Yadd = 0;

                Point1[0] = itemPosition.Key.X + 5;         //границы прямоугольника, в который должно попасть пересечение отрезков штампа (будущая т.вставки блока подписи)
                Point2[0] = itemPosition.Key.X + 24;

                Point1[1] = itemPosition.Key.Y - 4;
                Point2[1] = itemPosition.Key.Y;


                MySelection.Clear();
                short[] filterTypeLine = { 0 };
                object[] filterDataLine = { "LINE" };

                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                foreach (IAcadLine Line in ListLine)
                {
                    if ((Math.Round(Line.StartPoint[1]) == Math.Round(Line.EndPoint[1])) && ((Line.StartPoint[1] > Point1[1]) && (Line.StartPoint[1] < Point2[1])) && (((Line.StartPoint[0] < Point1[0]) && (Line.EndPoint[0] > Point2[0])) || (((Line.StartPoint[0] > Point1[0]) && (Line.EndPoint[0] < Point2[0])))))
                    {
                        if (Line.StartPoint[1] != Line.EndPoint[1])
                        {
                            Yadd = Math.Round(Line.StartPoint[1]);
                        }
                        else
                        {
                            Yadd = Line.StartPoint[1];
                        }
                    }
                    else if (((Math.Round(Line.StartPoint[0]) == Math.Round(Line.EndPoint[0])) && ((Line.StartPoint[0] > Point1[0]) && (Line.StartPoint[0] < Point2[0]))) && (((Line.StartPoint[1] < Point1[1]) && (Line.EndPoint[1] > Point2[1])) || (((Line.StartPoint[1] > Point1[1]) && (Line.EndPoint[1] < Point2[1])))))
                    {
                        if (Line.StartPoint[0] != Line.EndPoint[0])
                        {
                            Xadd = Math.Round(Line.StartPoint[0]);
                        }
                        else
                        {
                            Xadd = Line.StartPoint[0];
                        }
                    }
                }

                if ((Xadd != 0) && (Yadd != 0))
                {
                    string FileName = Blocks[itemPosition.Value];
                    double[] pt = new double[] { 0, 0, 0 };                     //точка вставки подписи

                    if (MaxY < Yadd)
                    {
                        MaxY = Yadd;
                    }
                    pt[0] = Xadd;
                    pt[1] = Yadd;

                    double[] pt1 = new double[] { 0, 0, 0 };

                    pt1[0] = pt[0] + 15;
                    pt1[1] = pt[1] + 1;
                    Point1 = new double[] { pt[0], pt[1], 0 };                  //границы поиска текста с датой
                    Point2 = new double[] { pt[0] + 25, pt[1] + 5, 0 };

                    MySelection.Clear();
                    filterDataLine[0] = "TEXT";
                    MySelection.Select(Mode: AcSelect.acSelectionSetWindow, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);

                    if (DateTimeRazrab != "")
                    {
                        foreach (AcadEntity item in MySelection)
                        {
                            item.Delete();                                      //удалить имеющуюся дату
                        }
                    }

                    Insert_Block_Podpis(pt, FileName, 0);
                    InsertDateTimeText(pt1, DateTimeRazrab, 0);
                }
                else
                {
                    Message.Add("Чертеж " + curDraw.Name + ": Не удалось вставить подпись для фамилии " + itemPosition.Value);
                }
            }
            ListExplodedObjects.Add(explodedObjects);
            return ListExplodedObjects;
        }

        /// <summary>
        /// Вычисление координат ближайших линий и вставка блока для левого(вертикального) штампа
        /// </summary>
        /// <param name="PosLeft"></param>
        /// <returns></returns>
        private List<dynamic> Insert_Podpis_With_Familiya_Left(SortedDictionary<Point, string> PosLeft)
        {
            short[] filterTypeLine = { 0 };
            object[] filterDataLine = { "LINE" };

            List<IAcadLine> ListLine = new List<IAcadLine>();
            List<dynamic> ListExplodedObjects = new List<dynamic>();

            if (LeftStampBlock == null)
            {
                Message.Add("В чертеже " + curDraw.Name + " отсутвствует форматная рамка");
                return ListExplodedObjects;
            }

            explodedObjects = LeftStampBlock.Explode();
            curDraw.Regen(AcRegenType.acAllViewports);

            if (explodedObjects != null)
            {
                foreach (AcadEntity item in explodedObjects)
                {
                    if (item.ObjectName == "AcDbLine")
                    {
                        ListLine.Add((IAcadLine)item);
                    }
                    if (item.ObjectName == "AcDbPolyline")
                    {
                        AcadLWPolyline LWPolyline = (AcadLWPolyline)item;

                        dynamic explodedObjects2 = null;
                        explodedObjects2 = LWPolyline.Explode();
                        foreach (AcadEntity item1 in explodedObjects2)
                        {
                            if (item1.ObjectName == "AcDbLine")
                            {
                                ListLine.Add((IAcadLine)item1);
                            }
                        }
                        ListExplodedObjects.Add(explodedObjects2);
                    }
                }
            }
            else
            {
                MySelection.Clear();
                filterDataLine[0] = "LWPOLYLINE";
                MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeLine, FilterData: filterDataLine);

                foreach (IAcadLWPolyline item in MySelection)
                {
                    dynamic explodedObjects2 = null;
                    explodedObjects2 = item.Explode();
                    ListExplodedObjects.Add(explodedObjects2);
                }
                curDraw.Regen(AcRegenType.acAllViewports);

                filterDataLine[0] = "LINE";
                MySelection.Clear();
                MySelection = curDraw.SelectionSets.Add("AddBlocksPodpis_VertLINE");
                MySelection.Select(Mode: AcSelect.acSelectionSetAll, FilterType: filterTypeLine, FilterData: filterDataLine);
                foreach (IAcadLine item in MySelection)
                {
                    ListLine.Add(item);
                }
            }

            double[] Point1 = new double[] { 0, 0, 0 };
            double[] Point2 = new double[] { 0, 0, 0 };
            double minpointaddX;
            double minpointaddY;
            double[] pt = new double[] { 0, 0, 0 };

            foreach (var itemPos in PosLeft)
            {
                double Xadd = 0;
                double Yadd = 0;
                minpointaddX = itemPos.Key.X;
                minpointaddY = itemPos.Key.Y;

                Point1[0] = minpointaddX;
                Point2[0] = minpointaddX + 4;

                Point1[1] = minpointaddY + 5;
                Point2[1] = minpointaddY + 24;

                foreach (IAcadLine Line in ListLine)
                {
                    //если координаты Y начала и конца отрезка совпадают
                    if (Math.Round(Line.StartPoint[1]) == Math.Round(Line.EndPoint[1]))
                    {
                        if ((Line.StartPoint[1] > Point1[1]) && (Line.StartPoint[1] < Point2[1]))
                        {
                            if ((Line.StartPoint[0] < Point1[0] && Line.EndPoint[0] > Point2[0]) || (Line.StartPoint[0] > Point1[0] && Line.EndPoint[0] < Point2[0]))
                            {
                                if (Line.StartPoint[1] != Line.EndPoint[1])
                                {
                                    Yadd = Math.Round(Line.StartPoint[1]);
                                }
                                else
                                {
                                    Yadd = Line.StartPoint[1];
                                }
                            }
                            if (((Line.StartPoint[0] < Point1[0]) && (Line.EndPoint[0] > Point1[0])) || ((Line.StartPoint[0] > Point2[0]) && (Line.EndPoint[0] < Point2[0])))
                            {
                                if (Line.StartPoint[1] != Line.EndPoint[1])
                                {
                                    Yadd = Math.Round(Line.StartPoint[1]);
                                }
                                else
                                {
                                    Yadd = Line.StartPoint[1];
                                }
                            }
                        }
                    }
                    else if (((Math.Round(Line.StartPoint[0]) == Math.Round(Line.EndPoint[0])) && ((Line.StartPoint[0] > Point1[0]) && (Line.StartPoint[0] < Point2[0]))) && (((Line.StartPoint[1] < Point1[1]) && (Line.EndPoint[1] > Point2[1])) || (((Line.StartPoint[1] > Point1[1]) && (Line.EndPoint[1] < Point2[1])))))
                    {
                        if (Line.StartPoint[0] != Line.EndPoint[0])
                        {
                            Xadd = Math.Round(Line.StartPoint[0]);
                        }
                        else
                        {
                            Xadd = Line.StartPoint[0];
                        }
                    }
                }
                if ((Xadd != 0) && (Yadd != 0))
                {
                    string FileName = BlocksLeft[itemPos.Value];

                    pt[0] = Xadd;
                    pt[1] = Yadd;

                    double[] pt1 = new double[] { 0, 0, 0 };
                    pt1[0] = pt[0] - 1;
                    pt1[1] = pt[1] + 15;
                    Point1 = new double[] { pt[0] - 5, pt[1] + 15, 0 };
                    Point2 = new double[] { pt[0], pt[1] + 25, 0 };

                    MySelection.Clear();
                    filterDataLine[0] = "TEXT";
                    MySelection.Select(Mode: AcSelect.acSelectionSetWindow, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                    filterDataLine[0] = "MTEXT";
                    MySelection.Select(Mode: AcSelect.acSelectionSetWindow, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);

                    if (DateTimeIzm != "")
                    {
                        foreach (AcadEntity item in MySelection)
                        {
                            item.Delete();
                        }
                    }

                    Insert_Block_Podpis(pt, FileName, 1.571);
                    InsertDateTimeText(pt1, DateTimeIzm, 1.571);
                }
                else
                {
                    Message.Add("Чертеж " + curDraw.Name + ": Не удалось вставить подпись для фамилии " + itemPos.Value);
                }
            }
            ListExplodedObjects.Add(explodedObjects);
            return ListExplodedObjects;
        }

        /// <summary>
        /// Вставить блок с подписью в указанную точку
        /// </summary>
        /// <param name="pt">Точка вставки</param>
        /// <param name="FileName">Путь к блоку подписи</param>
        private void Insert_Block_Podpis(double[] pt, string FileName, double rotation)
        {
            if (curDraw.ActiveSpace == AcActiveSpace.acPaperSpace)
            {
                curDraw.PaperSpace.InsertBlock(pt, FileName, 1, 1, 1, rotation);
            }
            else
            {
                curDraw.ModelSpace.InsertBlock(pt, FileName, 1, 1, 1, rotation);
            }
        }

        /// <summary>
        /// Вставка подписей в строки изменений
        /// </summary>
        /// <param name="PtIzmPodpis">Точка вставки нижней подписи в строках изменения</param>
        private bool Insert_Izm_Podpis_Right(double[] PtIzmPodpis)
        {
            //если найден штамп "ДУБЛИКАТ"
            if (StampDublicat != null)
            {
                dynamic varAttributes = StampDublicat.GetAttributes();

                foreach (AcadAttributeReference Att in varAttributes)  //найти в атрибутах штампа Дубликат фамилию
                {
                    if (Att.TagString == "FIO")
                    {
                        FIODublikat = Att.TextString;
                        break;
                    }
                }

                double[] pointDublikat = StampDublicat.InsertionPoint;
                pointDublikat[0] -= 32;                                     //смещение к точке вставки подписи
                pointDublikat[1] += 0;

                if (curDraw.ActiveSpace == AcActiveSpace.acPaperSpace)
                {
                    curDraw.PaperSpace.InsertBlock(pointDublikat, Blocks[FIODublikat], 1.3333, 1, 1, 0);
                }
                else
                {
                    curDraw.ModelSpace.InsertBlock(pointDublikat, Blocks[FIODublikat], 1.3333, 1, 1, 0);
                }
            }

            //Новый список для записи № изм. и точек вставки подписи
            CurrentIzmDict = new SortedDictionary<int, Izm>();
            double[] PtIzmPodpis1 = new double[] { 0, 0, 0 };
            double[] Point1;
            double[] Point2;
            short[] filterTypeLine = { 0 };
            object[] filterDataLine = { "TEXT" };

            //Собираем построчно имеющиеся номера изменений с точками вставки
            for (int i = 0; i < 4; i++)
            {
                //границы поиска текста столбца Изм.
                Point1 = new double[] { PtIzmPodpis[0] - 40, PtIzmPodpis[1] + 5 * i, 0 };                   
                Point2 = new double[] { PtIzmPodpis[0] - 30, PtIzmPodpis[1] + 5 * (i + 1), 0 };

                // Ищем текст в этих границах
                MySelection.Clear();
                filterDataLine[0] = "TEXT";
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing /*.acSelectionSetWindow*/, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                filterDataLine[0] = "MTEXT";
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing /*.acSelectionSetWindow*/, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                string izmNumber = "";

                //Если текст найден
                if (MySelection.Count > 0) 
                {
                    PtIzmPodpis1[1] = PtIzmPodpis[1] + 5 * i;
                    PtIzmPodpis1[0] = PtIzmPodpis[0];

                    foreach (AcadEntity item in MySelection)
                    {
                        switch (item.EntityName)
                        {
                            case "AcDbText":
                                AcadText acadText = (AcadText)item;
                                izmNumber = acadText.TextString;
                                if ( !CurrentIzmDict.ContainsKey(Convert.ToInt32(acadText.TextString.Trim())))
                                {
                                    CurrentIzmDict.Add(Convert.ToInt32(acadText.TextString.Trim()), new Izm(new double[]{ PtIzmPodpis1[0], PtIzmPodpis1[1], 0 }, "", DateTimeIzm));
                                }
                                break;
                            case "AcDbMText":
                                AcadMText acadMText = (AcadMText)item;
                                izmNumber = acadMText.TextString;
                                if ( !CurrentIzmDict.ContainsKey(Convert.ToInt32(acadMText.TextString.Trim())))
                                {
                                    CurrentIzmDict.Add(Convert.ToInt32(acadMText.TextString.Trim()), new Izm(new double[] { PtIzmPodpis1[0], PtIzmPodpis1[1], 0 }, "", DateTimeIzm));
                                }
                                break;
                        }
                    }
                }
                else break;

                Point1 = new double[] { PtIzmPodpis[0] - 20, PtIzmPodpis[1] + 5 * i, 0 };                   //границы поиска текста столбца Лист
                Point2 = new double[] { PtIzmPodpis[0] - 10, PtIzmPodpis[1] + 5 * (i + 1), 0 };

                MySelection.Clear();
                filterDataLine[0] = "TEXT";
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing /*.acSelectionSetWindow*/, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                filterDataLine[0] = "MTEXT";
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing /*.acSelectionSetWindow*/, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                string izmType = "";
                if (MySelection.Count > 0) 
                {
                    foreach (AcadEntity item in MySelection)
                    {
                        switch (item.EntityName)
                        {
                            case "AcDbText":
                                AcadText acadText = (AcadText)item;
                                izmType = acadText.TextString;
                                break;
                            case "AcDbMText":
                                AcadMText acadMText = (AcadMText)item;
                                izmType = acadMText.TextString;
                                break;
                        }
                    }
                }
                if (Convert.ToInt32(izmNumber) == CurrentIzmNumber && (izmType.StartsWith("Зам") || izmType.StartsWith("Нов")))
                {
                    ZamOrNov = true;
                }
            }

            if (CurrentIzmDict.Count > 0)
            {
                //Собираем в словарь отсутствующие в общем словаре номера изменений 
                Dictionary<int, Izm> tempDict = new Dictionary<int, Izm>();
                foreach (var item in CurrentIzmDict.Where(x => !CommonIzmDict.ContainsKey(x.Key)))
                {
                    tempDict.Add(item.Key, item.Value);
                }

                //Если есть отличия, запускаем форму
                if (tempDict.Count > 0 && !FlagAll)
                {
                    bool success = Configure_Request_Form(ref tempDict);
                    if (!success)
                    {
                        return false;
                    }
                    //Обновляем текущий словарь номеров изменений
                    foreach (var izm in tempDict)
                    {
                        CurrentIzmDict[izm.Key] = izm.Value;
                    }
                }

                if (FlagAll)
                {
                    foreach (var item in CurrentIzmDict)
                    {
                        item.Value.Date = DateTimeIzm;
                        item.Value.FIO = FIOIzm;
                    }
                }

                //Обновляем фамилии найденных номеров изменений в текущем словаре
                foreach (var item in CurrentIzmDict.Where(x => CommonIzmDict.ContainsKey(x.Key)))
                {
                    item.Value.FIO = CommonIzmDict[item.Key].FIO;
                    item.Value.Date = CommonIzmDict[item.Key].Date;
                }

                Insert_Some_Blocks_Podpis(CurrentIzmDict);
            }
            return true;
        }

        /// <summary>
        /// Цикличная вставка нескольких блоков подписей в строки изменений
        /// </summary>
        /// <param name="newDict">Найденные на чертеже номера изменений</param>
        private void Insert_Some_Blocks_Podpis(SortedDictionary<int, Izm> newDict)
        {
            foreach (var izm in newDict)
            {
                Insert_Block_Podpis(izm.Value.InsertionPoint, Blocks[izm.Value.FIO], 0);

                double[] Point1 = new double[] { izm.Value.InsertionPoint[0] + 15, izm.Value.InsertionPoint[1], 0 };                    //границы поиска текста с датой
                double[] Point2 = new double[] { izm.Value.InsertionPoint[0] + 25, izm.Value.InsertionPoint[1] + 5, 0 };

                short[] filterTypeLine = { 0 };
                object[] filterDataLine = { "TEXT" };

                MySelection.Clear();
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);
                filterDataLine[0] = "MTEXT";
                MySelection.Select(Mode: AcSelect.acSelectionSetCrossing /*.acSelectionSetWindow*/, Point1: Point1, Point2: Point2, FilterType: filterTypeLine, FilterData: filterDataLine);

                if (izm.Value.Date != "")
                {
                    foreach (AcadEntity item in MySelection)
                    {
                        item.Delete();
                    }
                }

                Point1[1] = Point1[1] + 1;
                InsertDateTimeText(Point1, izm.Value.Date, 0);
            }
        }
    }
}
