using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.ML;
using Microsoft.ML.Data;
using TestMLML.Model;
using Microsoft.ML.Trainers;
using Excel = Microsoft.Office.Interop.Excel;


namespace TestML
{
    public static class ProgramML
    {
        static void Main(string[] args)
        {
            
        }

        //------------------------------------------------------------------------------------------------------------------

        // Получение интерфейсно-независимых абсолютных адресов для сохранения файлов, получаемых в процессе работы программы 

        private static string path = AppDomain.CurrentDomain.BaseDirectory;

        private static string relPath1 = @"..\tempRentab.csv";
        private static string relPath2 = @"..\tempLikvid.csv";
        private static string relPath3 = @"..\tempDelActiv.csv";
        private static string relPath4 = @"..\tempFinUst.csv";

        private static string resPath1 = Path.Combine(path, relPath1);
        private static string resPath2 = Path.Combine(path, relPath2);
        private static string resPath3 = Path.Combine(path, relPath3);
        private static string resPath4 = Path.Combine(path, relPath4);

        private static string TRAIN_DATA_FILEPATH_1 = Path.GetFullPath(resPath1);
        private static string TRAIN_DATA_FILEPATH_2 = Path.GetFullPath(resPath2);
        private static string TRAIN_DATA_FILEPATH_3 = Path.GetFullPath(resPath3);
        private static string TRAIN_DATA_FILEPATH_4 = Path.GetFullPath(resPath4);

        private static string relModelPath1 = @"..\MLModel1.zip";
        private static string relModelPath2 = @"..\MLModel2.zip";
        private static string relModelPath3 = @"..\MLModel3.zip";
        private static string relModelPath4 = @"..\MLModel4.zip";

        private static string resModelPath1 = Path.Combine(path, relModelPath1);
        private static string resModelPath2 = Path.Combine(path, relModelPath2);
        private static string resModelPath3 = Path.Combine(path, relModelPath3);
        private static string resModelPath4 = Path.Combine(path, relModelPath4);

        private static string MODEL_FILEPATH_1 = Path.GetFullPath(resModelPath1);
        private static string MODEL_FILEPATH_2 = Path.GetFullPath(resModelPath2);
        private static string MODEL_FILEPATH_3 = Path.GetFullPath(resModelPath3);
        private static string MODEL_FILEPATH_4 = Path.GetFullPath(resModelPath4);

        // Инициализация базовых моделей машинного обучения

        private static MLContext mlContext1 = new MLContext(seed: 1);
        private static MLContext mlContext2 = new MLContext(seed: 1);
        private static MLContext mlContext3 = new MLContext(seed: 1);
        private static MLContext mlContext4 = new MLContext(seed: 1);

        //------------------------------------------------------------------------------------------------------------------

        // Структура выходных данных модели прогноза

        static ModelInput[] inputModelData = new ModelInput[]
        {
            new ModelInput
            {
                Id = 33
            },
            new ModelInput
            {
                Id = 34
            },
            new ModelInput
            {
                Id = 35
            },
            new ModelInput
            {
                Id = 36
            }
        };

        //------------------------------------------------------------------------------------------------------------------

        // Функция замены 

        private static string ChangeComma(string s)
        {
            return s.Replace(",", ".");
        }

        //------------------------------------------------------------------------------------------------------------------

        // Функция для работы с исходным файлом Excel, загрузка входных данных и их обработка

        public static void PrepareExcel(string excelPath, ref float[] arrRentab, ref float[] arrLikvid, ref float[] arrDelActiv, ref float[] arrFinUst)
        {
            // Создаем массивы для коэффициентов

            float[] arrKoeffTekusLikvid = new float[32];
            float[] arrKoeffAbsLikvid = new float[32];
            float[] arrKoeffSrochLikvid = new float[32];
            float[] arrKoeffRentActiv = new float[32];
            float[] arrKoeffRentSobstvKap = new float[32];
            float[] arrKoeffRentVlozhKap = new float[32];
            float[] arrKoeffOborachActiv = new float[32];
            float[] arrKoeffOborachSobstvKap = new float[32];
            float[] arrKoeffOborachVlozhKap = new float[32];
            float[] arrKoeffAvtonom = new float[32];
            float[] arrKoeffSootnZaemSobstvSred = new float[32];
            float[] arrKoeffManevrSobstOborSred = new float[32];

            // Создаем приложение для взаимодействия с Excel
            
            Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();

            // Открываем файл   
            
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(excelPath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

            // Создаем объект типа лист Excel 

            Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Считываем данные в массивы

            for (int numRow = 3; numRow < 15; numRow++)
            {
                Excel.Range usedRow = (Excel.Range)ObjWorkSheet.Cells[numRow, 1];
                
                if (usedRow.Value2 == "Коэффициент текущей (общей) ликвидности")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffTekusLikvid[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Value2 == "Коэффициент срочной ликвидности")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffSrochLikvid[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Value == "Коэффициент абсолютной ликвидности")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffAbsLikvid[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент рентабельности активов")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffRentActiv[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент рентабельности собств. капитала")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffRentSobstvKap[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент рентабельности вложенного капитала")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffRentVlozhKap[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент оборачиваемости активов")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffOborachActiv[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент оборачиваемости собств. капитала")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffOborachSobstvKap[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент оборачиваемости заемного капитала")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffOborachVlozhKap[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент автономии")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffAvtonom[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент соотношения заемных и собственных средств")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffSootnZaemSobstvSred[i - 2] = (float)localTemp.Value2;
                    }
                }

                else if (usedRow.Text == "Коэффициент маневренности собственных оборотных средств")
                {
                    for (int i = 2; i < 34; i++)
                    {
                        Excel.Range localTemp = (Excel.Range)ObjWorkSheet.Cells[numRow, i];
                        arrKoeffManevrSobstOborSred[i - 2] = (float)localTemp.Value2;
                    }
                }
            }

            // Заполнение конечных массивов и их запись в .csv файлы для последующего чтения фукнцией прогноза 

            for (int tempRentab = 0; tempRentab < 32; tempRentab++)
                arrRentab[tempRentab] = (float)Math.Round(Math.Sqrt(Math.Pow((double)arrKoeffRentActiv[tempRentab], 2) + Math.Pow((double)arrKoeffRentSobstvKap[tempRentab], 2) + Math.Pow((double)arrKoeffRentVlozhKap[tempRentab], 2)), 3, MidpointRounding.ToEven);
            for (int i = 0; i < 32; i++)
                arrRentab[i] *= 100;   // масштабирование значений для адекватного отображения на графике 
            System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_1, "id" + "," + "name" + "\r\n", Encoding.Unicode);
            for (int i = 1; i <= 32; i++)
            {
                System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_1, i.ToString() + "," + ChangeComma(arrRentab[i - 1].ToString()) + "\r\n", Encoding.Unicode);
            }

            for (int i = 0; i < 32; i++)
                arrLikvid[i] *= 100;
            for (int tempLikvid = 0; tempLikvid < 32; tempLikvid++)
                arrLikvid[tempLikvid] = (float)Math.Round(Math.Sqrt(Math.Pow((double)arrKoeffTekusLikvid[tempLikvid], 2) + Math.Pow((double)arrKoeffAbsLikvid[tempLikvid], 2) + Math.Pow((double)arrKoeffSrochLikvid[tempLikvid], 2)), 3, MidpointRounding.ToEven);
            System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_2, "id" + "," + "name" + "\r\n", Encoding.Unicode);
            for (int i = 1; i <= 32; i++)
            {
                System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_2, i.ToString() + "," + ChangeComma(arrLikvid[i - 1].ToString()) + "\r\n", Encoding.Unicode);
            }

            for (int tempDelActiv = 0; tempDelActiv < 32; tempDelActiv++)
                arrDelActiv[tempDelActiv] = (float)Math.Round(Math.Sqrt(Math.Pow((double)arrKoeffOborachActiv[tempDelActiv], 2) + Math.Pow((double)arrKoeffOborachSobstvKap[tempDelActiv], 2) + Math.Pow((double)arrKoeffOborachVlozhKap[tempDelActiv], 2)), 3, MidpointRounding.ToEven);
            System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_3, "id" + "," + "name" + "\r\n", Encoding.Unicode);
            for (int i = 1; i <= 32; i++)
            {
                System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_3, i.ToString() + "," + ChangeComma(arrDelActiv[i - 1].ToString()) + "\r\n", Encoding.Unicode);
            }

            for (int tempFinUst = 0; tempFinUst < 32; tempFinUst++)
                arrFinUst[tempFinUst] = (float)Math.Round(Math.Sqrt(Math.Pow((double)arrKoeffAvtonom[tempFinUst], 2) + Math.Pow((double)arrKoeffSootnZaemSobstvSred[tempFinUst], 2) + Math.Pow((double)arrKoeffManevrSobstOborSred[tempFinUst], 2)), 3, MidpointRounding.ToEven);
            System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_4, "id" + "," + "name" + "\r\n", Encoding.Unicode);
            for (int i = 1; i <= 32; i++)
            {
                System.IO.File.AppendAllText(TRAIN_DATA_FILEPATH_4, i.ToString() + "," + ChangeComma(arrFinUst[i - 1].ToString()) + "\r\n", Encoding.Unicode);
            }

        }

        //------------------------------------------------------------------------------------------------------------------

        // Вывод статистики по созданной модели

        private static void PrintStatistics(IEnumerable<TrainCatalogBase.CrossValidationResult<RegressionMetrics>> crossValidationResults)
        {
            var L1 = crossValidationResults.Select(r => r.Metrics.MeanAbsoluteError);
            var L2 = crossValidationResults.Select(r => r.Metrics.MeanSquaredError);
            var RMS = crossValidationResults.Select(r => r.Metrics.RootMeanSquaredError);
            var lossFunction = crossValidationResults.Select(r => r.Metrics.LossFunction);
            var R2 = crossValidationResults.Select(r => r.Metrics.RSquared);

            Console.WriteLine($"*************************************************************************************************************");
            Console.WriteLine($"*       Metrics for Regression model      ");
            Console.WriteLine($"*------------------------------------------------------------------------------------------------------------");
            Console.WriteLine($"*       Average L1 Loss:       {L1.Average():0.###} ");
            Console.WriteLine($"*       Average L2 Loss:       {L2.Average():0.###}  ");
            Console.WriteLine($"*       Average RMS:           {RMS.Average():0.###}  ");
            Console.WriteLine($"*       Average Loss Function: {lossFunction.Average():0.###}  ");
            Console.WriteLine($"*       Average R-squared:     {R2.Average():0.###}  ");
            Console.WriteLine($"*************************************************************************************************************");
        }

        //------------------------------------------------------------------------------------------------------------------

        // Получение конечного абсолютного адреса для сохранения модели

        private static string GetAbsolutePath(string relativePath)
        {
            FileInfo _dataRoot = new FileInfo(typeof(ProgramML).Assembly.Location);
            string assemblyFolderPath = _dataRoot.Directory.FullName;

            string fullPath = Path.Combine(assemblyFolderPath, relativePath);

            return fullPath;
        }

        //------------------------------------------------------------------------------------------------------------------

        // Основная функция по построению регрессионных моделей для каждого из массивов и по получению прогнозных значений

        public static void BuildPredictModel(ref float[] arrPredict1, ref float[] arrPredict2, ref float[] arrPredict3, ref float[] arrPredict4)
        {
            //Работа с первой моделью(первым массивом данных)

            //Подготовка данных для модели

            IDataView trainingDataView_1 = mlContext1.Data.LoadFromTextFile<ModelInput>(
                                            path: TRAIN_DATA_FILEPATH_1,
                                            hasHeader: true,
                                            separatorChar: ',',
                                            allowQuoting: true,
                                            allowSparse: false);

            IEstimator<ITransformer> dataProcessPipeline = mlContext1.Transforms.Concatenate("Features", new[] { "id" })
                                      .Append(mlContext1.Transforms.NormalizeMinMax("Features", "Features"))
                                      .AppendCacheCheckpoint(mlContext1);

            // Выбор тренировочного алгоритма

            var trainer = mlContext1.Regression.Trainers.LbfgsPoissonRegression(new LbfgsPoissonRegressionTrainer.Options() { DenseOptimizer = true, LabelColumnName = "name", FeatureColumnName = "Features" });         
            IEstimator<ITransformer> trainingPipeline = dataProcessPipeline.Append(trainer);

            // Тренировка модели

            ITransformer model1 = trainingPipeline.Fit(trainingDataView_1);

            // Оценка модели и вывод результатов оценки (для использования только в консоли)

            //var crossValidationResults = mlContext1.Regression.CrossValidate(trainingDataView_1, trainingPipeline, numberOfFolds: 5, labelColumnName: "name");
            //PrintStatistics(crossValidationResults);

            // Сохранение модели

            mlContext1.Model.Save(model1, trainingDataView_1.Schema, GetAbsolutePath(MODEL_FILEPATH_1));

            // Получение множественного прогноза (4 прогнозных значения)

            IDataView inputPredict = mlContext1.Data.LoadFromEnumerable<ModelInput>(inputModelData);
            IDataView predictions = model1.Transform(inputPredict);
            arrPredict1 = predictions.GetColumn<float>("Score").ToArray();

            //------------------------------------------------------------------------------------------------------------------

            // Работа со второй моделью (вторым массивом данных)

            // Подготовка данных для модели

            IDataView trainingDataView_2 = mlContext2.Data.LoadFromTextFile<ModelInput>(
                                            path: TRAIN_DATA_FILEPATH_2,
                                            hasHeader: true,
                                            separatorChar: ',',
                                            allowQuoting: true,
                                            allowSparse: false);

            IEstimator<ITransformer> dataProcessPipeline2 = mlContext2.Transforms.Concatenate("Features", new[] { "id" })
                                      .Append(mlContext2.Transforms.NormalizeMinMax("Features", "Features"))
                                      .AppendCacheCheckpoint(mlContext2);

            // Выбор тренировочного алгоритма

            var trainer2 = mlContext2.Regression.Trainers.LbfgsPoissonRegression(new LbfgsPoissonRegressionTrainer.Options() { DenseOptimizer = true, LabelColumnName = "name", FeatureColumnName = "Features" });
            IEstimator<ITransformer> trainingPipeline2 = dataProcessPipeline2.Append(trainer2);

            // Тренировка модели

            ITransformer model2 = trainingPipeline2.Fit(trainingDataView_2);

            // Оценка модели и вывод результатов оценки (для использования только в консоли)

            //var crossValidationResults2 = mlContext2.Regression.CrossValidate(trainingDataView_2, trainingPipeline2, numberOfFolds: 5, labelColumnName: "name");
            //PrintStatistics(crossValidationResults2);

            // Сохранение модели

            mlContext2.Model.Save(model2, trainingDataView_2.Schema, GetAbsolutePath(MODEL_FILEPATH_2));

            // Получение множественного прогноза (4 прогнозных значения)

            IDataView inputPredict2 = mlContext2.Data.LoadFromEnumerable<ModelInput>(inputModelData);
            IDataView predictions2 = model2.Transform(inputPredict2);
            arrPredict2 = predictions2.GetColumn<float>("Score").ToArray();

            //------------------------------------------------------------------------------------------------------------------

            // Работа с третьей моделью (третьим массивом данных)

            // Подготовка данных для модели

            IDataView trainingDataView_3 = mlContext3.Data.LoadFromTextFile<ModelInput>(
                                            path: TRAIN_DATA_FILEPATH_3,
                                            hasHeader: true,
                                            separatorChar: ',',
                                            allowQuoting: true,
                                            allowSparse: false);

            IEstimator<ITransformer> dataProcessPipeline3 = mlContext3.Transforms.Concatenate("Features", new[] { "id" })
                                      .Append(mlContext3.Transforms.NormalizeMinMax("Features", "Features"))
                                      .AppendCacheCheckpoint(mlContext3);

            // Выбор тренировочного алгоритма

            var trainer3 = mlContext3.Regression.Trainers.LbfgsPoissonRegression(new LbfgsPoissonRegressionTrainer.Options() { DenseOptimizer = true, LabelColumnName = "name", FeatureColumnName = "Features" });
            IEstimator<ITransformer> trainingPipeline3 = dataProcessPipeline3.Append(trainer3);

            // Тренировка модели

            ITransformer model3 = trainingPipeline3.Fit(trainingDataView_3);

            // Оценка модели и вывод результатов оценки (для использования только в консоли)

            //var crossValidationResults3 = mlContext3.Regression.CrossValidate(trainingDataView_3, trainingPipeline3, numberOfFolds: 5, labelColumnName: "name");
            //PrintStatistics(crossValidationResults3);

            // Сохранение модели

            mlContext3.Model.Save(model3, trainingDataView_3.Schema, GetAbsolutePath(MODEL_FILEPATH_3));

            // Получение множественного прогноза (4 прогнозных значения)

            IDataView inputPredict3 = mlContext3.Data.LoadFromEnumerable<ModelInput>(inputModelData);
            IDataView predictions3 = model3.Transform(inputPredict3);
            arrPredict3 = predictions3.GetColumn<float>("Score").ToArray();

            //------------------------------------------------------------------------------------------------------------------

            // Работа с четвертой моделью (четвертым массивом данных)

            // Подготовка данных для модели

            IDataView trainingDataView_4 = mlContext4.Data.LoadFromTextFile<ModelInput>(
                                            path: TRAIN_DATA_FILEPATH_4,
                                            hasHeader: true,
                                            separatorChar: ',',
                                            allowQuoting: true,
                                            allowSparse: false);

            IEstimator<ITransformer> dataProcessPipeline4 = mlContext4.Transforms.Concatenate("Features", new[] { "id" })
                                      .Append(mlContext4.Transforms.NormalizeMinMax("Features", "Features"))
                                      .AppendCacheCheckpoint(mlContext4);

            // Выбор тренировочного алгоритма

            var trainer4 = mlContext4.Regression.Trainers.LbfgsPoissonRegression(new LbfgsPoissonRegressionTrainer.Options() { DenseOptimizer = true, LabelColumnName = "name", FeatureColumnName = "Features" });
            IEstimator<ITransformer> trainingPipeline4 = dataProcessPipeline4.Append(trainer4);

            // Тренировка модели

            ITransformer model4 = trainingPipeline4.Fit(trainingDataView_4);

            // Оценка модели и вывод результатов оценки (для использования только в консоли)

            //var crossValidationResults4 = mlContext4.Regression.CrossValidate(trainingDataView_4, trainingPipeline4, numberOfFolds: 5, labelColumnName: "name");
            //PrintStatistics(crossValidationResults4);

            // Сохранение модели

            mlContext4.Model.Save(model4, trainingDataView_4.Schema, GetAbsolutePath(MODEL_FILEPATH_4));

            // Получение множественного прогноза (4 прогнозных значения)

            IDataView inputPredict4 = mlContext4.Data.LoadFromEnumerable<ModelInput>(inputModelData);
            IDataView predictions4 = model4.Transform(inputPredict4);
            arrPredict4 = predictions4.GetColumn<float>("Score").ToArray();
        }


    }
};

