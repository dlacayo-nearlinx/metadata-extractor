using System.Text;

namespace MetadataExtractor.Common.Models
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
}