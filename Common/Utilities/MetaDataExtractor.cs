using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Text;
using CodeFluent.Runtime.BinaryServices;
using MetadataExtractor.Common.Models;
using MetadataExtractor.Common.Extensions;

namespace MetadataExtractor.Common.Utilities
{

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

       
        /// <summary>
        /// Reads the file info and metadata for the specified file
        /// </summary>
        /// <param name="file"></param>
        /// <returns>FileMetadata object with values</returns>
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
        private static string GetErrorMessage(string property)
        {
            return string.Format("The {0} information could not be read", property);
        }

        private static string GetKeywordsUsingCodeFluent(FileInfo file)
        {
            const int PropertyCountOffset = 52; 
            const int PropertyStart = 56;
            const int KeywordPropertyId = 5;
            

            var streams = NtfsAlternateStream.EnumerateStreams(file);
            var sb = new StringBuilder();
            foreach (var stream in streams)
            {
                var type = stream.StreamType;
                
                if (type == NtfsAlternateStreamType.AlternateData && stream.Name.Contains("SummaryInformation") && !stream.Name.Contains("Document"))
                {
                    var ads = file.FullName + stream.Name;
                    var bytes = NtfsAlternateStream.ReadAllBytes(ads);

                    // get the number of props
                    var propCount = Convert.ToInt32(bytes[PropertyCountOffset]);
                    for (int i = 0; i < propCount; i++) 
                    { 
                        var index = PropertyStart + (8*i);
                        var propId = Convert.ToInt32(bytes[index]);
                        if (propId != KeywordPropertyId )
                            continue;

                        // we have the prop, advance 4 bytes and get the offset, need to shift bits because of Big Endian                       
                        var offset = bytes[index+5] << 8 | bytes[index + 4];

                        // header data offset + value offset + padding bytes
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
