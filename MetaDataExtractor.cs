using CodeFluent.Runtime.BinaryServices;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Trinet.Core.IO.Ntfs;

namespace WindowsFormsADSTest
{

    public class FileMetadata 
    {
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Author { get; set; }
        public string Category { get; set; }        
        public string Comments { get; set; }
        //
        public string Name { get; set; }
        public string Extension { get; set; }
        public string Folder { get; set; }
        public string Path { get; set; }        
        public string Created { get; set; }
        public string Modified { get; set; }
        public string Size { get; set; }
        //
        public string Keywords { get; set; }

        public override string ToString()
        {
            var sb = new StringBuilder();

            sb.AppendLine(string.Format("Title:\t {0}", Title));
            sb.AppendLine(string.Format("Subject:\t {0}", Subject));
            sb.AppendLine(string.Format("Author:\t {0}", Author));
            sb.AppendLine(string.Format("Category:\t {0}", Category));
            sb.AppendLine(string.Format("Comments:\t {0}", Comments));
            sb.AppendLine(string.Format("Name:\t {0}", Name));
            sb.AppendLine(string.Format("Extension:\t {0}", Extension));
            sb.AppendLine(string.Format("Folder:\t {0}", Folder));
            sb.AppendLine(string.Format("Path:\t {0}", Path));
            sb.AppendLine(string.Format("Created:\t {0}", Created));
            sb.AppendLine(string.Format("Modified:\t {0}", Modified));
            sb.AppendLine(string.Format("Size:\t {0}", Size));
            sb.AppendLine(string.Format("Keywords:\t {0}", Keywords));
            

            return sb.ToString();
        }
    }

     public static class DictionaryExtensions 
    {
         public static TValue GetValueOrDefault<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue = default(TValue))
         {
             if (dictionary == null) return defaultValue;
             if (key == null) return defaultValue;

             TValue value;
             return dictionary.TryGetValue(key, out value) ? value : defaultValue;
         }  
    }

    /*
     Alternate data streams.
     * IDT File Format with ascii data
     */   

    public static class MetaDataExtractor
    {       
        const string Title = "Title";
        const string Subject = "Subject";
        const string Author = "Author";
        const string Category = "Category";
        const string Comments = "Comments";
        const string Name = "Name";
        const string Extension = "Extension";
        const string Folder = "Folder";
        const string Path = "Path";
        const string Created = "Date Created";
        const string Modified = "Date Modified";
        const string Size = "Size";
        const string Keywords = "Keywords";

        private static string GetErrorMessage(string property) 
        {
            return string.Format("The {0} information could not be read", property);
        }

        public static FileMetadata GetFileMetadata(FileInfo file)
        {
            var properties = GetExtendedPropertiesFromFolderData(file);
            var keywords = GetKeywordsUsingCodeFluent(file);
            
            var metadata = new FileMetadata() 
            {
                Title = properties.GetValueOrDefault(Title, GetErrorMessage(Title)),
                Subject = properties.GetValueOrDefault(Subject, GetErrorMessage(Subject)),
                Author = properties.GetValueOrDefault(Author, GetErrorMessage(Author)),
                Category = properties.GetValueOrDefault(Category, GetErrorMessage(Category)),
                Comments = properties.GetValueOrDefault(Comments, GetErrorMessage(Comments)),
                //
                Name = properties.GetValueOrDefault(Name, GetErrorMessage(Name)),
                Extension = properties.GetValueOrDefault(Extension, GetErrorMessage(Extension)),
                Folder = properties.GetValueOrDefault(Folder, GetErrorMessage(Folder)),
                Path = properties.GetValueOrDefault(Path, GetErrorMessage(Path)),
                Created = properties.GetValueOrDefault(Created, GetErrorMessage(Created)),
                Modified = properties.GetValueOrDefault(Modified, GetErrorMessage(Modified)),
                Size = properties.GetValueOrDefault(Size, GetErrorMessage(Size)),
                //
                Keywords = keywords
            };
            return metadata;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private static string GetExtendedPropertiesFromFile(FileInfo file)
        {
            const string ExtendedPropertyGuid = "{F29F85E0-4FF9-1068-AB91-08002B27B3D9}";
            var sb = new StringBuilder();            
            
            Shell32.Shell shell = new Shell32.Shell();
            Shell32.Folder directory = shell.NameSpace(file.DirectoryName);
            Shell32.FolderItem folderItem = directory.Items().Item(file.Name);
            Shell32.ShellFolderItem shellFolderItem = (Shell32.ShellFolderItem)folderItem;


            string name0 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 0");
            string name1 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 1");
            string title = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 2");
            string subject = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 3");
            string author = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 4");
            string keywords = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 5");
            string comments = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 6");
            string name7 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 7");
            string name8 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 8");
            string name9 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 9");
            string name10 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 10");
            string name11 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 11");
            string name12 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 12");
            string name13 = (string)shellFolderItem.ExtendedProperty(ExtendedPropertyGuid + " 13");


            return sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        private static string GetKeywordsUsingCoreIO(FileInfo file)
        {
            //FileInfo file = new FileInfo(path.FullName);
            var sb = new StringBuilder();

            sb.AppendLine(string.Format("Processing file: {0}", file.Name));
            // List the additional streams for a file:
            sb.AppendLine("Processing Alternate Data Streams for file");
            foreach (AlternateDataStreamInfo stream in file.ListAlternateDataStreams())
            {
                //stream.
                //Console.WriteLine("{0} - {1} bytes", s.Name, s.Size);
                sb.AppendLine(string.Format("{0} - {1} bytes", stream.Name, stream.Size));
                var atts = stream.Attributes;
                var fileStream = stream.OpenRead();
                var canRead = fileStream.CanRead;
                var streamReader = new StreamReader(fileStream, true);
                var adStream = streamReader.ReadToEnd();
                sb.AppendLine("Stream Contents:\n\n");
                //sb.AppendLine(adStream);
                var contentsReader = stream.OpenText();
                var encoding = contentsReader.CurrentEncoding;
                var contents = contentsReader.ReadToEnd();
            }

            // Read the "Zone.Identifier" stream, if it exists:
            if (file.AlternateDataStreamExists("Zone.Identifier"))
            {
                //Console.WriteLine("Found zone identifier stream:");
                sb.AppendLine("Found zone identifier stream:");

                AlternateDataStreamInfo s = file.GetAlternateDataStream("Zone.Identifier", FileMode.Open);
                using (TextReader reader = s.OpenText())
                {
                    var ss = reader.ReadToEnd();
                    Console.WriteLine(ss);
                    sb.AppendLine(ss);
                }
                // Delete the stream:
                s.Delete();
            }
            else
            {
                //Console.WriteLine("No zone identifier stream found.");
                sb.AppendLine("No zone identifier stream found.");
            }

            // Alternative method to delete the stream:
            file.DeleteAlternateDataStream("Zone.Identifier");
            var text = sb.ToString();
            return text;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public static string GetKeywordsUsingCodeFluent(FileInfo file)
        {
            var streams = NtfsAlternateStream.EnumerateStreams(file);
            foreach (var stream in streams)
            {
                var name = stream.Name;
                var size = stream.Size;
                var type = stream.StreamType;

                if (type == NtfsAlternateStreamType.AlternateData && stream.Name.Contains("SummaryInformation") && !stream.Name.Contains("Document"))
                {
                    var ads = file.FullName + stream.Name;
                    var bytes = NtfsAlternateStream.ReadAllBytes(ads);
                                        
                    var testFile = file.FullName + ".streamdata";
                    File.WriteAllBytes(testFile,bytes);
                    
                    var asciiData = System.Text.Encoding.ASCII.GetString(bytes);
                    var testTextFile = file.FullName + ".streamdata" + ".txt";
                    System.IO.File.WriteAllText(testTextFile,asciiData);
                }
            }

            return "";
        }

        public static Dictionary<string, string> GetExtendedPropertiesFromFolderData(FileInfo file)
        {
            var filePath = file.FullName;
            var directory = file.DirectoryName;
            var extension = file.Extension;
            var shell = new Shell32.Shell();
            var folder = shell.NameSpace(directory);
            var fileName = file.Name;
            var folderitem = folder.ParseName(fileName);
            var dictionary = new Dictionary<string, string>();
            var i = -1;
            while (++i < 320)
            {
                var header = folder.GetDetailsOf(null, i);
                if (String.IsNullOrEmpty(header)) continue;
                var value = folder.GetDetailsOf(folderitem, i);
                if (!dictionary.ContainsKey(header)) dictionary.Add(header, value);                
            }
            if (!dictionary.ContainsKey("Folder")) dictionary.Add("Folder", directory);
            if (!dictionary.ContainsKey("Path")) dictionary.Add("Path", filePath);
            if (!dictionary.ContainsKey("Extension")) dictionary.Add("Extension", extension);

            return dictionary;
        }
    }
}
