using SolidWorks.Interop.dsgnchk;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


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



    // Основной класс библиотеки
    public class SWDrawer
    {



        public SldWorks app;               // Экземпляр приложения Solidworks
        public IModelDoc2 model;           // Активный документ (модель)
        public PartDoc part;

        public SketchManager skMng;        // Менеджер скечей (предоставляет функционал для скетчей)
        public FeatureManager ftMng;       // Менеджер фич(??) (нужен для создания бобышек и отверстий)
        public SelectionMgr selMng;









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
                    MessageBox.Show("Хуета, не могу создать проект такого типа!");
                    break;
            }

            model = (IModelDoc2)app.IActiveDoc2;
            part = (PartDoc)model;

            skMng = model.SketchManager;
            ftMng = model.FeatureManager;
            selMng = model.SelectionManager;
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
        // Возвращает массив элементов - стороны и диагонали
        public SketchSegment[] createRectangle(double bX, double bY, double bZ, double eX, double eY, double eZ)
        {
            return (SketchSegment[])skMng.CreateCornerRectangle(bX, bY, bZ, eX, eY, eZ);
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


        // Отладочный метод, создающий куб со стороной 1 метр 
        public void fastCube()
        {
            model.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, false, 0, null, 0);
            skMng.InsertSketch(true);

            // Рисуем квадрат 1x1 метр (1000 мм)
            skMng.CreateCenterRectangle(0, 0, 0, 0.5, 0.5, 0); // центр в начале координат, половина стороны

            //skMng.InsertSketch(false); // завершаем эскиз

            // Выдавливаем куб на 1 метр
            ftMng.FeatureExtrusion2(
                true, false, false,
                (int)swEndConditions_e.swEndCondBlind,
                0, 1.0, 0, false, false,
                false, false, 0, 0, false, false, false, false,
                true, true, true,
                0, 0, false
            );
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

        // Метод, выделяющий грань тела по индексу массива его граней
        public bool SelectFaceByIndex(Body2 body, int faceIndex)
        {
            Face2[] faces = getAllFaces(body);
            Face2 face = faces[faceIndex];

            return ((Entity)face).Select2(false,0);
        }

        
        // Метод, поочередно показывающий все грани тела и указывающий их индекс в массиве граней тела
        public void viewBodyFaces(Body2 body)
        {
            Face2[] faces = getAllFaces(body);

            for (int j = 0; j < faces.Length; j++)
            {
                SelectFaceByIndex(body, j);
                app.SendMsgToUser("Тело: " + body.Name + "\n" + "Грань: " + j);

            }
        }

    }
}
