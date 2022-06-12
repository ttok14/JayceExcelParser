using System;
using System.IO;
using McMaster.Extensions.CommandLineUtils;
using Serilog;
using JayceExcelParser;
using JayceExcelParser.Generate;
using JayceExcelParser.Common;
using Serilog.Sinks.SystemConsole.Themes;

/// <summary>
/// ========================================================
/// Packages
///     - CommandLineUtils 
///         [Doc https://natemcmaster.github.io/CommandLineUtils/]
///         [Github https://github.com/natemcmaster/CommandLineUtils]
///     - Serilog
///         [Doc ]
///         [Github ]
/// ========================================================
/// Configuration [<see cref="Configuration"/>]
///     - 개발용 환경 설정 클래스
///         - e.g. 실제 / 시뮬레이션 모드 [<see cref="Configuration.CurrentMode"/>]
/// ========================================================
/// </summary>
namespace JayceExcelParser
{
    class Program
    {
        static int Main(string[] args)
        {
            Configuration.Setup(ref args);

            //==============================================//

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console(theme: SystemConsoleTheme.Colored)
                .CreateLogger();
            Log.Information("Excel Parser Tool initiating . . .");

            var app = new CommandLineApplication();

            app.HelpOption();

            /// 참고 
            ///         - 다음 타입 <see cref="CommandOptionType.SingleOrNoValue"/> 은 Cmd 랑 Value 사이에 Space 사용 불가
            ///                 e.g. --name Banana  
            ///                             => Error 
            ///                        --name:Banana
            ///                             =>  OK 
            ///                        --name=Banana
            ///                             => OK
            ///                        --name
            var optionExcelDir = app.Option("-xl|--excel <EXCEL_LOCATION>", "Directory of excel files, current directory if empty", CommandOptionType.SingleValue);
            optionExcelDir.DefaultValue = Directory.GetCurrentDirectory();
            var optionSaveDir = app.Option("-s|--save <SAVE_LOCATION>", "Directory of the output, current directory if empty", CommandOptionType.SingleValue);
            optionSaveDir.DefaultValue = Directory.GetCurrentDirectory();

            app.OnExecute(() =>
            {
                var result = default(ExitCode);
                var generator = new Generator();
                var desc = new GeneratorDesc(optionExcelDir.Value(), optionSaveDir.Value());

                if (Configuration.CurrentMode == Mode.SIMULATION)
                {
                    #region ====:: Test Code ::====
                    desc.AddMessagePack();
                    #endregion
                }

                result = Helper.GenerateResultToExitCode(generator.Generate(desc));
                EmitExitCodeMsg(result);
                return (int)result;
            });

            return app.Execute(args);
        }

        static void EmitExitCodeMsg(ExitCode code)
        {
            if (code == ExitCode.SUCCESS)
            {
                Log.Information($"Successfuly finished");
            }
            else
            {
                Log.Error($"Error occured : {code}");
            }
        }
    }
}
