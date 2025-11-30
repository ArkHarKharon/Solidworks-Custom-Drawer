using SolidWorks.Interop.dsgnchk;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;


namespace Temp
{
    /*
           Данная библиотека пытается собрать все самые используемые методы Solidworks API
           в одном месте, тем самым упрощая процесс создание моделей, а также помогая
           в изучении API.

           Автор предполагает, что пользователи приложения будут вводить размеры в миллиметрах,
           после чего размеры будут сразу переводиться в метры и сохраняться в переменные. Так
           можно будет не переводить единицы измерения в каждом методе проекта.

           Автор рекомендует выключить привязки в настройках Solidworks: Параметры (шестеренка) ->
           Эскиз -> Взаимосвязи/привязки -> выключить "Разрешить привязки"
           
       */



    // Перечисление с типами документов Solidworks
    public enum DocumentType
    {
        DRAWING,
        PART,
        ASSEMBLY
    }

    // Перечисление имен начальных поверхностей
    public enum DefaultPlaneName
    {
        TOP,
        FRONT,
        RIGHT,
        TEST
    }

    // перечисление типов вырезания отверстия
    public enum HoleType
    {
        CUT_THROUGH = swEndConditions_e.swEndCondThroughAll,
        DISTANCE = swEndConditions_e.swEndCondBlind,
        OFFSET = swEndConditions_e.swEndCondOffsetFromSurface
    }

    // Перечисление с основными типами объектов
    public enum ObjectType
    {
        EDGE = swSelectType_e.swSelEDGES,
        FACE = swSelectType_e.swSelFACES,
        VERTEX = swSelectType_e.swSelVERTICES,
        PLANE = swSelectType_e.swSelDATUMPLANES,
        AXIS = swSelectType_e.swSelDATUMAXES,
        SKETCH_SEGMENT = swSelectType_e.swSelSKETCHSEGS
    }

    // Основной класс библиотеки
    public class SWDrawer
    {
        public SldWorks app;               // Экземпляр приложения Solidworks
        public ModelDoc2 model;           // Активный документ (модель)
        public PartDoc part;               // Документ детали
        public AssemblyDoc assembly;       // Документ сборки

        public SketchManager skMng;        // Менеджер скечей (предоставляет функционал для скетчей)
        public FeatureManager ftMng;       // Менеджер фич(??) (нужен для создания бобышек и отверстий)
        public SelectionMgr selMng;        // Менеджер выбора элементов 









        /*
                    --- Методы инициализации проекта ---
        
            Далее приведены методы, необходимые для иницализации класса, запуска Solidworks
            и создания проекта.
         
        */

        // Пытается подключиться к открытому экземпляру Soldiworks
        // Если не получается - капускает новый экземпляр
        public bool init()
        {
            // Открыть SolidWorks либо получить экземпляр открытого приложения
            try
            {
                // Попытка получить существующий экземпляр SolidWorks
                app = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                return true;
            }
            catch
            {
                try
                {
                    // Если не нашли открытый - создаем новый
                    app = new SldWorks();
                    app.FrameState = (int)swWindowState_e.swWindowMaximized;
                    app.Visible = true;
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось подключиться к SolidWorks: {ex.Message}");
                    return false;
                }
            }
        }


        // Создает новый проект, тип проекта выбирается по перечислению DocumentType
        // Принимает необязательный параметр - номер шаблона чертежа
        public void newProject(DocumentType documentType, int drawingTemplateNum = 10)
        {
            // Проверяем, что SolidWorks инициализирован
            if (app == null)
            {
                MessageBox.Show("SolidWorks не инициализирован. Вызовите init() сначала.");
                return;
            }

            switch (documentType)
            {
                case DocumentType.DRAWING:
                    app.NewDrawing(drawingTemplateNum);
                    break;

                case DocumentType.PART:
                    app.NewPart();
                    break;

                case DocumentType.ASSEMBLY:
                    app.NewAssembly();
                    break;

                default:
                    MessageBox.Show("Не могу создать проект такого типа!");
                    break;
            }

            // Получаем активный документ
            model = (ModelDoc2)app.IActiveDoc2;

            if (model == null)
            {
                MessageBox.Show("Не удалось создать документ");
                return;
            }

            // Размеры в миллиметрах
            model.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);

            // Приводим к PartDoc только если это деталь
            if (model is PartDoc)
                part = (PartDoc)model;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;
        }

        // Пытается подключиться к открытой детали
        public bool connectToOpenedPart()
        {
           

            // Получаем активный документ
            model = app.IActiveDoc2 as ModelDoc2;

            if (model == null)
            {
                app.SendMsgToUser("Нет активного документа в SolidWorks.");
                return false;
            }

            // Проверяем что документ — деталь
            if (!(model is PartDoc))
            {
                app.SendMsgToUser("Активный документ не является деталью (PartDoc).");
                return false;
            }

            // Привязка PartDoc
            part = (PartDoc)model;

            // Инициализация менеджеров
            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;

            model.SetUnits(
                (short)swLengthUnit_e.swMM,
                (short)swFractionDisplay_e.swDECIMAL,
                0, 0, false
            );

            app.SendMsgToUser("Подключен к открытой детали.");
            return true;
        }


        // Пытается подключиться к открытой сборке
        public bool connectToOpenedAssembly()
        {
            // Получаем активный документ
            model = app.IActiveDoc2 as ModelDoc2;

            if (model == null)
            {
                app.SendMsgToUser("Нет активного документа в SolidWorks.");
                return false;
            }

            // Проверяем что документ — деталь
            if (!(model is AssemblyDoc))
            {
                app.SendMsgToUser("Активный документ не является сборкой (AssemblyDoc).");
                return false;
            }

            // Привязка PartDoc
            assembly = (AssemblyDoc)model;

            // Инициализация менеджеров
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;

            model.SetUnits(
                (short)swLengthUnit_e.swMM,
                (short)swFractionDisplay_e.swDECIMAL,
                0, 0, false
            );

            app.SendMsgToUser("Подключен к открытой сборке.");
            return true;
        }






        /*
                        --- Методы для работы с эскизами ---
         
            Далее приведены методы для работы с эскизами и чертежами. Большиство методов (особенно
            методы создания элеметов имеют тип сохраняемого значения SketchSegment. Для последующей
            работы с этими элементами рекомендую сохранять их в переменные. Сделать это можно
            следующим образом:


            SWDrawer drawer = new Drawer;

                ...

            SketchSegment smallCircle = drawer.createCircle (...);

            Подобным образом сохранять в переменные необходимо все элементы, которые будут использо-
            ваться далее, например, для обрезания.
         
        */


        // Создание прямой по кооринатам 2 точек - начальной и конечной
        public SketchSegment createLineByCoords(double bX, double bY, double bZ, double eX, double eY, double eZ)
        {
            return skMng.CreateLine(bX, bY, bZ, eX, eY, eZ);
        }


        // Создание окружности по координатам центра и точки на окружности
        public SketchSegment createCircleByPoint(double cX, double cY, double cZ, double rX, double rY, double rZ)
        {
            return skMng.CreateCircle(cX, cY, cZ, rX, rY, rZ);
        }


        // Создание окружности по координатам центра и радиуса
        public SketchSegment createCircleByRadius(double cX, double cY, double cZ, double radius)
        {
            return skMng.CreateCircleByRadius(cX, cY, cZ, radius);
        }


        // Создает прямоугольник по координатам 2 противолежащих вершин
        public void createRectangle(double bX, double bY, double bZ, double eX, double eY, double eZ)
        {
            skMng.CreateCornerRectangle(bX, bY, bZ, eX, eY, eZ);
        }


        // Создает прямоугольник по координатам 2 точек - точки пересечения диагоналей и вершине
        public void createCenterRectangle(double cX, double cY, double cZ, double eX, double eY, double eZ)
        {
            skMng.CreateCenterRectangle(cX, cY, cZ, eX, eY, eZ);
        }


        // Метод обрезает элемент эскиза
        // Принимает обраезаемый элемент и примерные координаты места разреза
        // Дополнительным параметром является режим обрезания (все режимы можно узнать на
        // сайте Solidworks)
        public void trim(SketchSegment segment, double X, double Y, double Z, int option = 0)
        {
            //Очищает выбор элементов эскиза
            model.ClearSelection2(true);

            segment.Select(true);
            skMng.SketchTrim(option, X, Y, Z);
            segment.Select(false);
        }








        /*
                    --- Методы для работы с деталями --- 
            
            Далее приведены основные методы для создания деталей: выбор плоскостей, вставка эскизов,
            вытягивания, вырезание отверстий и т. п.
         
        */


        // Метод позволяет выделить начальную плоскость (Спереди, Справа и Сверху)
        // Принимает название плоскости из перечисления DefaultPlaneName
        public void selectDefaultPlane(DefaultPlaneName planeName)
        {
            switch (planeName)
            {
                case DefaultPlaneName.TOP:
                    model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;

                case DefaultPlaneName.FRONT:
                    model.Extension.SelectByID2("СПЕРЕДИ", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;

                case DefaultPlaneName.RIGHT:
                    model.Extension.SelectByID2("СПРАВА", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("Right Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;

                default:
                    app.SendMsgToUser("Не удалось получить плоскость " + planeName.ToString());
                    break;
            }
        }


        // Метод, открывающий/закрывающий скетч на выделенной плоскости или грани
        public void insertSketch(bool start)
        {
            skMng.InsertSketch(start);
        }

        public void selectSketchByNumber(int number)
        {
            model.ClearSelection2(true);
            model.Extension.SelectByID2("Sketch" + number, "SKETCH", 0, 0, 0, false, 0, null, 0);
            model.Extension.SelectByID2("Эскиз" + number, "SKETCH", 0, 0, 0, false, 0, null, 0);


        }


        // Отладочный метод, создающий куб со стороной 1 метр 
        public void fastCube()
        {
            selectDefaultPlane(DefaultPlaneName.TOP);
            insertSketch(true);

            createCenterRectangle(0, 0, 0, 0.5, 0.5, 0);

            extrude(1);

        }


        // Метод, возвращающий массив тел проекта
        public Body2[] getAllBodies()
        {
            if (part == null) return new Body2[0];

            object bodiesObj = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
            if (bodiesObj == null) return new Body2[0];

            object[] objArray = bodiesObj as object[];
            if (objArray == null || objArray.Length == 0) return new Body2[0];

            // Приведение каждого элемента к Body2
            return objArray.Cast<Body2>().ToArray();
        }


        // Метод, возвращающий массив граней тела
        public Face2[] getAllFaces(Body2 body)
        {
            object[] facesObj = body.GetFaces();

            object[] objArray = facesObj as object[];

            return objArray.Cast<Face2>().ToArray();
        }

        public Edge[] getAllEdges(Face2 face)
        {
            object[] edgesObj = face.GetEdges();

            object[] objArray = edgesObj as object[];

            return objArray.Cast<Edge>().ToArray();
        }


        // Метод, выделяющий грань тела по индексу массива его граней
        public void SelectFaceByIndex(Body2 body, int faceIndex)
        {
            Face2[] faces = getAllFaces(body);
            Face2 face = faces[faceIndex];

            Entity faceEnt = (Entity)face;
            faceEnt.Select2(false, 0);
        }

        //Метод, выбирающий ребро по индексу массива ребер грани
        public void SelectEdgeByIndex(Face2 face, int edgeIndex)
        {
            Edge[] edges = getAllEdges(face);

            Entity edge = (Entity)edges[edgeIndex];
            edge.Select2(false, 0);
        }


        // Метод, поочередно показывающий все грани тела и указывающий их индекс в массиве граней тела
        // Нужен для определения индекса нужной грани и последующего использования в SelectFaceByIndex()
        public void viewBodyFaces(Body2 body)
        {
            Face2[] faces = getAllFaces(body);

            for (int j = 0; j < faces.Length; j++)
            {
                SelectFaceByIndex(body, j);
                app.RunCommand(169, ""); //int для комманды поворота к нормали
                app.SendMsgToUser("Тело: " + body.Name + "\n" + "Грань: " + j);

            }

        }


        // Метод, поочередно показывающий все ребра грани и указывающий их индекс в массиве ребер грани
        // Нужен для определения индекса нужного ребра и последующего использования в SelectEdgeByIndex()
        public void viewFaceEdges(int faceNumber)
        {
            Body2[] bodies = getAllBodies();
            Face2 face = getAllFaces(bodies[0])[faceNumber];

            SelectFaceByIndex(bodies[0], faceNumber);
            app.RunCommand(169, ""); //int для комманды поворота к нормали

            model.ClearSelection2(true);

            Edge[] edges = getAllEdges(face);

            for (int i = 0; i < edges.Length; i++)
            {
                Entity edge = (Entity)edges[i];
                edge.Select2(false, 0);
                app.SendMsgToUser("Грань: " + faceNumber + "\n" + "Ребро: " + i);

            }

        }


        // Вытягивает скетч, который был вставлен (важно не выходить из скетча для работы метода)
        // Принимает длину вытягивания
        // Если вытягивание происходит не в ту сторону, поставить changeDirection = true
        public void extrude(double extrusionLength, bool changeDirection = false)
        {
            ftMng.FeatureExtrusion2(
               true, false, changeDirection,
               (int)swEndConditions_e.swEndCondBlind,
               0, extrusionLength, 0, false, false,
               false, false, 0, 0, false, false, false, false,
               true, true, true,
               0, 0, false
           );
        }




        // Вырезает отверстие по скетчу, который был вставлен (важно не выходить из скетча для работы метода)
        // Принимает обязательный параметр - тип вырезания по enum HoleType
        // Также принимает необязательные параметры - длину выреза и флаг для смены направления выреза
        public void cutHole(HoleType typeOfHole, bool changeDirection = false, double depth = 0)
        {
            ftMng.FeatureCut4(
                true, false, changeDirection,
                (int)typeOfHole, 0,
                depth, 0.0,
                false, false, false, false,
                0.0, 0.0,
                false, false, false, false, false,
                true, true, true, true, false,
                0, 0, false, false
            );
        }


        // Метод создает скругление на выбранных ребрах
        // Принимает номер грани faceNum, на которой находятся скругляемые ребра, номер брать из viewBodyFaces()
        // Принимает радиус скругления, а также номера рёбер, полученные из viewFaceEdges()
        public void createFillets(int faceNum, double radius, params int[] edgeNums)
        {
            model.ClearSelection2(true);

            Body2[] bodies = getAllBodies();
            Face2 face = getAllFaces(bodies[0])[faceNum];
            Edge[] edges = getAllEdges(face);

            for (int i = 0; i < edgeNums.Length; i++)
            {
                Entity edgeEnt = (Entity)edges[edgeNums[i]];
                edgeEnt.Select2(true, 0);
            }

            ftMng.FeatureFillet3(
                         2,
                         radius,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0,
                         0
                        );
        }


        // Метод создает фаски на выбранных ребрах
        // Принимает номер грани faceNum, на которой находятся необходимые ребра, номер брать из viewBodyFaces()
        // Принимает 2 значения double - расстояния от ребра до начала фасок, а также номера рёбер, полученные из viewFaceEdges()
        public void createChamfers(int faceNum, double distance1, double distance2, params int[] edgeNums)
        {
            model.ClearSelection2(true);

            Body2[] bodies = getAllBodies();
            Face2 face = getAllFaces(bodies[0])[faceNum];
            Edge[] edges = getAllEdges(face);

            for (int i = 0; i < edgeNums.Length; i++)
            {
                Entity edgeEnt = (Entity)edges[edgeNums[i]];
                edgeEnt.Select2(true, 0);
            }

            ftMng.InsertFeatureChamfer((int)swFeatureChamferOption_e.swFeatureChamferTangentPropagation,
                (int)swChamferType_e.swChamferDistanceDistance, distance1, 0, distance2, 0, 0, 0);


        }


        // Метод для удаления нескольких элементов дерева детали, которые были созданы последними
        // Принимает количество эементов с конца, которые необходимо удалить
        public void deleteLastFeatures(int numberOfElements)
        {
            if (model == null || numberOfElements <= 0)
                return;

            List<Feature> featureList = new List<Feature>();
            Feature feat = model.FirstFeature();

            // Собираем все фичи в список
            while (feat != null)
            {
                featureList.Add(feat);
                feat = feat.GetNextFeature();
            }

            // Удаляем последние n фич
            int count = featureList.Count;
            for (int i = count - 1; i >= count - numberOfElements && i >= 0; i--)
            {
                Feature f = featureList[i];
                if (f != null)
                {
                    bool selected = f.Select2(false, -1);
                    if (selected)
                    {
                        model.Extension.DeleteSelection2((int)swDeleteSelectionOptions_e.swDelete_Absorbed);
                    }
                }
            }
        }




        /*
                   --- Методы для работы со сборками --- 

           Далее приведены основные методы для раборты со сборками: добавление деталей, (хз, не придумал пока...)

       */

        // Метод импортирует деталь в сборку
        // Принимает абсолютный путь до файла .SLDPRT или .SLDASM, также принимает x,y и z примерного местоположения

        public void importPart(string path, double x, double y, double z)
        {
            model = (ModelDoc2)app.IActiveDoc2;
            assembly = model as AssemblyDoc;

            app.OpenDoc6(
                      path,
                      Path.GetExtension(path).ToLower() == ".sldasm"
                          ? (int)swDocumentTypes_e.swDocASSEMBLY
                          : (int)swDocumentTypes_e.swDocPART,
                      (int)swOpenDocOptions_e.swOpenDocOptions_LoadLightweight,
                      "", 0, 0);

            assembly.AddComponent2(path, x, y, z);

            app.CloseDoc(path);

        }


        // Импортирует все детали из папки
        // Принимает абсолютный путь до папки, а также расстояние между центрами деталей
        public void importPartsFromFolder(string folderPath, double offset = 0.1, int rawLength = 5)
        {
            String[] files = Directory.GetFiles(folderPath);

            double x = 0;
            double y = 0;

            for (int i = 0; i < files.Length; i++)
            {
                importPart(files[i], x , 0, y);

                x += offset;

                if (i % rawLength == 0 && i != 0)
                {
                    y += offset;
                    x = 0;
                }
                    
            }
        }



        // Метод возвращает список тел сборки
        public List<Body2> GetAllBodiesFromAssembly()
        {
            var result = new List<Body2>();

            if (assembly == null)
                return result;

            // Получаем все компоненты первого уровня
            object[] comps = (object[])assembly.GetComponents(true); // true = рекурсивно

            foreach (object o in comps)
            {
                Component2 comp = o as Component2;
                if (comp == null) continue;

                ModelDoc2 model = comp.GetModelDoc2();
                if (model == null) continue;

                PartDoc part = model as PartDoc;
                if (part == null) continue; // компонент может быть подпроектирован

                // Получаем тела у детали
                object bodiesObj = part.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodiesObj == null) continue;

                foreach (Body2 b in (object[])bodiesObj)
                {
                    result.Add(b);
                }
            }

            return result;
        }



        /*
                         --- Методы для обработки выбора пользователя (бета) --- 

            Далее приведены основные методы для обработки выбора ползователя

        */

        // Метод, позволяющий получить массив выделенных объектов
        public List<object> GetSelectedObjects()
        {
            List<object> result = new List<object>();

            if (model == null)
                return result;

            int count = selMng.GetSelectedObjectCount();

            for (int i = 1; i <= count; i++)
            {
                object obj = selMng.GetSelectedObject6(i, -1);
                if (obj != null)
                    result.Add(obj);
            }

            return result;
        }



    }

    /*
                ---- Вспомогательный класс для вырезания массива отверстий -----

        Класс предоставляет базовый функционал, необходимый для вырезания отверсий в теле. Место вырезания 
        отверстий ограничивается объемом, полученным 2 выделенными точками.
     
     
     */

    enum Directions
    {
        OX,
        OY,
        OZ
    }


    public class HolesArrayCutter
    {
        public class Point
        {
            public double x;
            public double y;
            public double z;

            public Point() { }

            public Point(double x, double y, double z)
            {
                this.x = x;
                this.y = y;
                this.z = z;
            }

            public void set(double x, double y, double z)
            {
                this.x = x;
                this.y = y;
                this.z = z;
            }
        }


        SWDrawer drawer = new SWDrawer();

        Point bottomLeftPoint = new Point();
        Point topRightPoint = new Point();



        private int verticiesNum;
        private List<Point> circlesCenters;
        private double radius;

        private double offsetX;
        private double offsetY;

        int verticalDirection;

        public HolesArrayCutter(SWDrawer drawer)
        {
            this.drawer = drawer;
        }


        // Основной метод для вырезания массива отверстий
        // Принимает число отверстий в ряду и строке, количество вершин у эскиза отверстия
        // offset позволяет сделать отверстие внутри тела, на расстоянии offset от нижей и верхней граней области
        // Для работы метода необходимо предварительно выделить 2 вершины (Vertex)
        // Эти вершины образуют призматическую область, в котороый вырезаются отверстия
        public void cutHoles(int rowsNum, int columsNum, int verticiesNum,double angle = 0.0, double offset = 0.0, bool reverse = false)
        {
            List<object> temp = drawer.GetSelectedObjects();
            List<Vertex> selected = new List<Vertex>();

            foreach (object obj in temp)
            {
                if (obj is Vertex)
                    selected.Add((Vertex)obj);
            }


            if (selected == null)
            {
                drawer.app.SendMsgToUser("Не удалось получить массив выбранных элементов!");
                return;
            }
                

            Vertex bottom = (Vertex)selected[0];
            Vertex top = (Vertex)selected[1];

            setPoints(bottom, top);
            setCircles(rowsNum, columsNum);

            if (offset == 0.0)
            {
                drawer.selectDefaultPlane(DefaultPlaneName.TOP);
            }

            else
            {
                // Выбор для 1 ограничения
                drawer.model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 0, null, 0);
                drawer.model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 0, null, 0);


                // Выбор для 2 ограничения
                drawer.model.Extension.SelectByID2("СВЕРХУ", "PLANE", 0, 0, 0, false, 1, null, 0);
                drawer.model.Extension.SelectByID2("Top Plane", "PLANE", 0, 0, 0, false, 1, null, 0);

                RefPlane planeRef =  drawer.ftMng.InsertRefPlane((int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Parallel, 0.0,
                    (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Distance, offset, 0, 0.0);

                Entity plane = (Entity)planeRef;

                drawer.model.ClearSelection2(true);

                plane.Select(true);
            }

            drawer.insertSketch(true);

            foreach (Point p in circlesCenters)
            {
                drawer.skMng.CreatePolygon(p.x, -p.z, 0, p.x + radius * Math.Cos(angle * Math.PI / 180.0), -p.z + radius * Math.Sin(angle * Math.PI / 180.0), 0, verticiesNum, false);
            }

            if (offset == 0.0)
                drawer.cutHole(HoleType.CUT_THROUGH, true);
            else
            {
                double depth = Math.Abs(top.GetPoint()[2] - bottom.GetPoint()[2]) - 2 * offset;
                drawer.app.SendMsgToUser("Глубина:" +  depth);
                drawer.cutHole(HoleType.DISTANCE, true, depth);

            }
        }
        

        //Метод, обрабатывающий выделение точек пользователем
        private void setPoints(Vertex bottom, Vertex top)
        {
            bottomLeftPoint.set(bottom.GetPoint()[0], bottom.GetPoint()[1], bottom.GetPoint()[2]);
            topRightPoint.set(top.GetPoint()[0], top.GetPoint()[1], top.GetPoint()[2]);

         
        }

        // Метод определяет местоположение центров окружностей и их радиус
        // В эти окружности вписываются эскизы отверстий
        private void setCircles(int rowsNum, int columsNum, double offset = 0)
        {
            if (rowsNum <= 0) throw new ArgumentException("rowsNum must be > 0");
            if (columsNum <= 0) throw new ArgumentException("columsNum must be > 0");

            circlesCenters = new List<Point>();

            // Координаты прямоугольника
            double minX = Math.Min(bottomLeftPoint.x, topRightPoint.x);
            double maxX = Math.Max(bottomLeftPoint.x, topRightPoint.x);
            double minZ = Math.Min(bottomLeftPoint.z, topRightPoint.z);
            double maxZ = Math.Max(bottomLeftPoint.z, topRightPoint.z);

            double totalSpaceX = maxX - minX;
            double totalSpaceZ = maxZ - minZ;

            if (totalSpaceX <= 0 || totalSpaceZ <= 0)
                throw new InvalidOperationException("Invalid rectangle size.");

            // Минимальный зазор
            double minGap = offset > 0 ? offset : 0.02 * Math.Min(totalSpaceX, totalSpaceZ);
            if (minGap < 1e-9) minGap = 1e-9;

            // --- Вычисление максимально возможных радиусов по каждой оси ---
            double radiusX = (totalSpaceX - (columsNum + 1) * minGap) / (2.0 * columsNum);
            double radiusZ = (totalSpaceZ - (rowsNum + 1) * minGap) / (2.0 * rowsNum);

            if (radiusX <= 0 || radiusZ <= 0)
                throw new InvalidOperationException("Not enough space for circles");

            // Итоговый радиус — минимальный (чтобы точно влезало по обеим сторонам)
            radius = Math.Min(radiusX, radiusZ);

            // Перерасчёт отступов по итоговому радиусу
            offsetX = (totalSpaceX - 2.0 * radius * columsNum) / (columsNum + 1);
            offsetY = (totalSpaceZ - 2.0 * radius * rowsNum) / (rowsNum + 1);

            if (offsetX < 0) offsetX = 0;
            if (offsetY < 0) offsetY = 0;

            // Стартовые координаты
            double startX = minX + offsetX + radius;
            double startZ = minZ + offsetY + radius;

            // --- Формирование сетки окружностей ---
            for (int r = 0; r < rowsNum; r++)
            {
                for (int c = 0; c < columsNum; c++)
                {
                    double cx = startX + c * (2.0 * radius + offsetX);
                    double cz = startZ + r * (2.0 * radius + offsetY);

                    circlesCenters.Add(new Point(cx, bottomLeftPoint.y, cz));
                }
            }
        }



    }
}
