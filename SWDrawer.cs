using SolidWorks.Interop.dsgnchk;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;


namespace Lab_5
{
    /*
           Чуточку переделанная библиотека от GitHub: @ArkHarKharon
            
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
    }



    // Основной класс библиотеки
    public class SWDrawer
    {



        public SldWorks app;               // Экземпляр приложения Solidworks
        public IModelDoc2 model;           // Активный документ (модель)
        public PartDoc part;

        public SketchManager skMng;        // Менеджер скечей (предоставляет функционал для скетчей)
        public FeatureManager ftMng;       // Менеджер фич(??) (нужен для создания бобышек и отверстий)
        public SelectionMgr selMng;




        // Extra methods by KingArtur1000

        /// <summary>
        /// Сбрасывает модель: либо создаёт новый документ, либо очищает текущий.
        /// </summary>
        /// <param name="newDoc">true = создать новый Part, false = очистить текущий</param>
        public void ResetModel(bool newDoc = false)
        {
            if (newDoc)
            {
                // Создать новый Part-документ
                app.NewPart();
            }
            else
            {
                app.CloseAllDocuments(true);
                app.NewPart();
            }

            model = (IModelDoc2)app.IActiveDoc2;
            part = (PartDoc)model;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = (SelectionMgr)model.SelectionManager;
        }


        /// <summary>
        /// Создаёт луч (задаёт направление через rX, rY, rZ), если на пересечении с указанной точкой находит грань - выбираем её
        /// </summary>
        public bool SelectFaceByRayMM(double wX, double wY, double wZ, double rX, double rY, double rZ)
        {
            if (model == null) return false;

            bool isFaceSelected = model.Extension.SelectByRay(wX / 1000, wY / 1000, wZ / 1000, rX / 1000, rY / 1000, rZ / 1000, 0.01, 2, false, 0, 0);

            if (isFaceSelected) { return true; }
            else { MessageBox.Show("Не удалось выбрать грань!"); return false; }
        }




        /*
                    --- Методы инициализации проекта ---
        
            Далее приведены методы, необходимые для иницализации класса, запуска Solidworks
            и создания проекта.
         
            Данный модуль нуждается в переработке / улучшении.
        */






        // Закрывает текущие экземпляры Solidworks и запускает новый 
        public void init()
        {
            Process[] processes = Process.GetProcessesByName("SLDWORKS");
            foreach (Process process in processes)
            {
                process.CloseMainWindow();
                process.Kill();
            }

            app = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
            app.Visible = true;


        }


        // Создает новый проект, тип проекта выбирается по перечислению DocumentType
        // Принимает необязательный параметр - номер шаблона чертежа
        public void newProject(DocumentType documentType, int drawingTemplateNum = 10)
        {
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
                    MessageBox.Show("Что-то не то... Не могу создать проект такого типа!");
                    break;
            }

            model = (IModelDoc2)app.IActiveDoc2;
            part = (PartDoc)model;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = (SelectionMgr)model.SelectionManager;
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
                    model.Extension.SelectByID2("TOP", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;

                case DefaultPlaneName.FRONT:
                    model.Extension.SelectByID2("СПЕРЕДИ", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("FRONT", "PLANE", 0, 0, 0, false, 0, null, 0);
                    break;

                case DefaultPlaneName.RIGHT:
                    model.Extension.SelectByID2("СПРАВА", "PLANE", 0, 0, 0, false, 0, null, 0);
                    model.Extension.SelectByID2("RIGHT", "PLANE", 0, 0, 0, false, 0, null, 0);
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
            object[] facesObj = (object[])body.GetFaces();

            object[] objArray = facesObj as object[];

            return objArray.Cast<Face2>().ToArray();
        }

        // Метод, выделяющий грань тела по индексу массива его граней
        public bool SelectFaceByIndex(Body2 body, int faceIndex)
        {
            Face2[] faces = getAllFaces(body);
            Face2 face = faces[faceIndex];

            return ((Entity)face).Select2(false, 0);
        }


        // Метод, поочередно показывающий все грани тела и указывающий их индекс в массиве граней тела
        // Нужен для определения индекса нужной грани и последующего использования в SelectFaceByIndex()
        public void viewBodyFaces(Body2 body)
        {
            Face2[] faces = getAllFaces(body);

            for (int j = 0; j < faces.Length; j++)
            {
                SelectFaceByIndex(body, j);
                app.SendMsgToUser("Тело: " + body.Name + "\n" + "Грань: " + j);

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

    }
}