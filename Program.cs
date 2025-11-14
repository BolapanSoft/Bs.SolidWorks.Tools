using Bs.SolidWorks.Tools.Commands;
using Bs.SolidWorks.Tools.Logging;
using SolidWorks.Interop.sldworks;
using System;
using System.IO;

namespace Bs.SolidWorks.Tools {
    internal class Program {
        static void Main(string[] args) {
            try {
                if (args.Length == 0) { 
                    Console.WriteLine("❌ Не указаны параметры командной строки."); 
                    Console.WriteLine("Использование: Bs.SolidWorks.Tools --ExtractConfigurations");
                    return; 
                }

                foreach (var item in args) {
                    if (item.StartsWith("--")) {
                        // Обработка параметров командной строки
                        switch (item) {
                            case "--ExtractConfigurations":
                                ExtractConfigurations();
                                break;
                            default:
                                Console.WriteLine("❌ Неизвестный параметр: " + item);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex) {
                Console.WriteLine("Ошибка: " + ex.Message);
                Console.ReadKey();
            }
        }
        static void ExtractConfigurations() {
            var currentDirectory = Directory.GetCurrentDirectory();
            // Реализация извлечения конфигураций
            var logger = new LightLogger(currentDirectory, "Bs.ExtractConfigurations.log", maxFileBytes: 5 * 1024 * 1024, maxFiles: 5, externalWriter:Console.Out);
            logger.Level = LogLevel.Debug;
            SldWorks? swApp = null;
            try {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var assemblyName = assembly.GetName();
                string version = assembly.GetName().Version.ToString(3);
                logger.Info($"{assemblyName} v{version}");
                swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks;
                if (swApp == null) {
                    logger.Error("Не удалось подключиться к SolidWorks.");
                    return;
                }
                swApp.Visible = false;
                logger.Info("✅ Подключено к SolidWorks " + swApp.RevisionNumber());
                logger.Info("Запуск извлечения конфигураций...");
                // Логика извлечения конфигураций
                var job=new ExtractConfigurationsJob(swApp, logger);
                job.ExtractAll(currentDirectory);
                logger.Info("Извлечение конфигураций завершено успешно.");
            }
            catch (Exception ex) {
                logger.Error("Ошибка при извлечении конфигураций: " + ex.Message);
            }
            finally {
                if (swApp != null) {
                    swApp.CloseAllDocuments(true);
                    swApp.ExitApp();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(swApp);
                    swApp = null;
                }
                logger.Dispose();
            }
        }   
    }

}
