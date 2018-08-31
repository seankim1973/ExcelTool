using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace ExcelTool
{
    public class Utils
    {
        static DirectoryInfo _outputDir = null;
        public static DirectoryInfo OutputDir
        {
            get
            {
                return _outputDir;
            }
            set
            {
                _outputDir = value;
                if (!_outputDir.Exists)
                {
                    _outputDir.Create();
                }
            }
        }
        public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(OutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
        public static FileInfo GetFileInfo(DirectoryInfo altOutputDir, string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(altOutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }

        internal static DirectoryInfo GetDirectoryInfo(string directory)
        {
            var di = new DirectoryInfo(_outputDir.FullName + Path.DirectorySeparatorChar + directory);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }

        public static DirectoryInfo GetCodeBasePath(bool forOutputDir = true)
        {
            Directory.SetCurrentDirectory(Directory.GetParent(TestContext.CurrentContext.TestDirectory).ToString());
            string baseDir = Directory.GetParent(Directory.GetCurrentDirectory()).ToString();

            var directory = (forOutputDir) ? (Directory.GetParent(baseDir)).GetDirectories("TestCases")[0] : (Directory.GetParent(baseDir));
            return directory;
        }
    }

    public static class EnumHelper
    {
        public static int GetIndex(this Enum value)
        {
            int output;

            try
            {
                Type type = value.GetType();
                FieldInfo fi = type.GetField(value.ToString());
                IndexValue[] attrs = fi.GetCustomAttributes(typeof(IndexValue), false) as IndexValue[];
                output = attrs[0].Value;
            }
            catch (Exception)
            {
                throw;
            }
            return output;
        }
    }

    public class IndexValue : Attribute
    {
        public IndexValue(int value)
        {
            Value = value;
        }
        public int Value { get; }
    }
}
