using DSOFile;
using System;
using System.Collections.Generic;

namespace MetadataFile
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = args.Length > 0 ? args[0] : @"C:\Users\David\Configuration.db";

            OleDocumentPropertiesClass file = new OleDocumentPropertiesClass();

            Console.WriteLine("Opening file: {0}", filePath);
            file.Open(filePath, false, dsoFileOpenOptions.dsoOptionDefault);

            SetCustomProperty(file.CustomProperties, "Packets", 65);
            SetCustomProperty(file.CustomProperties, "Powered by", "Rey David Dominguez");

            foreach (CustomProperty p in file.CustomProperties)
            {
                Console.WriteLine("{0}:{1}", p.Name, p.get_Value().ToString());
            }

            file.Close(true);
            Console.ReadKey();
        }
        
        public static void SetCustomProperty(CustomProperties customProperties, string property, object value)
        {
            foreach(CustomProperty customProperty in customProperties)
            {
                if(customProperty.Name == property)
                {
                    customProperty.set_Value(ref value);
                    return;
                }
            }

            customProperties.Add(property, ref value);
        }
    }
}
