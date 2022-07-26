using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using CodeFluent.Runtime.BinaryServices;
using MetadataExtractor.Common.Models;
using MetadataExtractor.Common.Extensions;

namespace MetadataExtractor.Common.Utilities
{
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

        // not used
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

        private static string GetKeywordsUsingCodeFluent(FileInfo file)
        {
            var streams = NtfsAlternateStream.EnumerateStreams(file);
            var sb = new StringBuilder();
            foreach (var stream in streams)
            {
                var name = stream.Name;
                var size = stream.Size;
                var type = stream.StreamType;
                
                if (type == NtfsAlternateStreamType.AlternateData && stream.Name.Contains("SummaryInformation") && !stream.Name.Contains("Document"))
                {
                    var ads = file.FullName + stream.Name;
                    var bytes = NtfsAlternateStream.ReadAllBytes(ads);
                      
                    /*
                    // save stram data to raw file
                    var testFile = file.FullName + ".streamdata";
                    File.WriteAllBytes(testFile,bytes);                    
                    */

                    // get the number of props
                    var propCount = Convert.ToInt32(bytes[52]);
                    for (int i = 0; i < propCount; i++) 
                    { 
                        var index = 56 + (8*i);
                        var propId = Convert.ToInt32(bytes[index]);
                        if (propId != 5)
                            continue;

                        // we have the prop, advance 4 bytes and get the offset                        
                        var offset = bytes[index+5] << 8 | bytes[index + 4];

                        var dataLocation = 0x2C + offset + 0xC;
                        var b = bytes[dataLocation];
                        while ( b != 0 && b != 30) 
                        {
                            sb.Append((char)b);
                            dataLocation++;
                            b = bytes[dataLocation];
                        }                        
                        break;
                    }                                     
                }
            }
            var kw = sb.ToString();
            return kw;
        }

        private static Dictionary<string, string> GetExtendedPropertiesFromFolderData(FileInfo file)
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
