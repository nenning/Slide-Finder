using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Configuration;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using XnaFan.ImageComparison;
using System.Reflection;
using System.IO.Compression;
using System.Diagnostics;
using System.Drawing;

namespace ppt
{

    class Program
    {     
        static void Main(string[] args)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            string masterSlides = ConfigurationManager.AppSettings["masterSlides"];
            string targetSlides = ConfigurationManager.AppSettings["yourSlides"];
            if (args.Length > 0)
            {
                targetSlides = args[0];
                Console.WriteLine("using slides from " + targetSlides);
            }
            Console.WriteLine($"Master slides: {masterSlides}");
            Console.WriteLine($"Your slides: {targetSlides}");

            string currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            Directory.SetCurrentDirectory(currentDirectory);
            var tempDirMaster = Path.Combine(currentDirectory, "master");
            var tempDirTarget = Path.Combine(currentDirectory, "target");
            string resultFile = Path.Combine(currentDirectory, "results.txt");
            string masterInfo = Path.Combine(currentDirectory, "master.info");
            if (File.Exists(resultFile))
            {
                File.Delete(resultFile);
            }

            // don't recreate master thumbnails if they exist and source file hasn't changed
            bool createMasterThumbs = true;
            var modifiedDate = File.GetLastWriteTimeUtc(masterSlides);
            if (File.Exists(masterInfo))
            {
                var masterData = File.ReadAllText(masterInfo);
                File.Delete(masterInfo);
                createMasterThumbs = (masterData != masterSlides + modifiedDate) || !Directory.Exists(tempDirMaster) || !Directory.EnumerateFiles(tempDirMaster).Any();
            }

            if (createMasterThumbs && Directory.Exists(tempDirMaster))
            {
                Directory.Delete(tempDirMaster, true);
            }
            if (Directory.Exists(tempDirTarget))
            {
                Directory.Delete(tempDirTarget, true);
            }
            Directory.CreateDirectory(tempDirMaster);
            Directory.CreateDirectory(tempDirTarget);

            // TODO use real examples

            FindSourceSlides(masterSlides, targetSlides, resultFile, tempDirMaster, tempDirTarget, createMasterThumbs);

            Process.Start("notepad.exe", resultFile);
            File.WriteAllText(masterInfo, masterSlides + modifiedDate);
            Directory.Delete(tempDirTarget, true);
            
            stopwatch.Stop();
            Console.WriteLine("Elapsed={0}", stopwatch.Elapsed);
            return;
        }

        private static void FindSourceSlides(string masterSlides, string targetSlides, string resultFile, string tempDirMaster, string tempDirTarget, bool createMasterThumbs)
        {
            var masterTexts = ExtractText(masterSlides);
            var targetTexts = ExtractText(targetSlides);

            var result = new SlideReference[targetTexts.Count + 1];
            for (int i = 0; i < result.Length; i++)
            {
                result[i] = new SlideReference(i);
            }

            Parallel.Invoke(
                () => CompareImages(masterSlides, targetSlides, tempDirMaster, tempDirTarget, result, createMasterThumbs),
                () => CompareTexts(masterTexts, targetTexts, result)
            );

            var sb = new StringBuilder();
            sb.AppendLine($"Which slide of '{targetSlides}' matches to which slide of '{masterSlides}'");
            foreach (var elem in result.OrderBy(e => e.SourceSlide))
            {
                if (elem.SourceSlide == 0) continue;
                sb.AppendLine(elem.ToString());
            }
            File.WriteAllText(resultFile, sb.ToString());
        }

        private static void CompareTexts(IDictionary<int, string> masterTexts, IDictionary<int, string> targetTexts, SlideReference[] result)
        {
            int count = 0;
            Parallel.ForEach(targetTexts, (targetSlide) =>
            {
                var entry = result[targetSlide.Key];
                Interlocked.Increment(ref count);
                Console.WriteLine($"comparing text {count} of {targetTexts.Count-1}");

                foreach (var sourceSlide in masterTexts)
                {
                    var distance = StringDistance.Compute(targetSlide.Value, sourceSlide.Value);
                    entry.UpdateText(sourceSlide.Key, distance);
                }
            });
        }

        private static void CompareImages(string masterSlides, string targetSlides, string tempDirMaster, string tempDirTarget, SlideReference[] result, bool createMasterThumbs)
        {
            Parallel.Invoke(
                () => { if (createMasterThumbs) SaveSlidesAsImages(masterSlides, tempDirMaster); } ,
                () => SaveSlidesAsImages(targetSlides, tempDirTarget)
             );

            CompareThumbnails(tempDirMaster, tempDirTarget, result);
        }

        private static void CompareThumbnails(string masterDir, string targetDir, SlideReference[] result)
        {
            int count = 0;
            var targetSlides = Directory.GetFiles(targetDir);
            Parallel.ForEach(targetSlides, (targetFile) =>
            {
                var targetSlideNr = GetSlideNumber(targetFile);
                var entry = result[targetSlideNr];

                Interlocked.Increment(ref count);
                Console.WriteLine($"comparing image {count} of {result.Length-1}");

                var masterSlides = Directory.GetFiles(masterDir);
                //Parallel.ForEach (masterSlides, (masterFile) => 
                foreach (var masterFile in masterSlides)
                {
                    int masterSlideNr = GetSlideNumber(masterFile);
                    var diff = ImageTool.GetPercentageDifference(masterFile, targetFile);
                    entry.UpdateImage(masterSlideNr, diff);
                } //);
            });
        }

        private static int GetSlideNumber(string file)
        {
            var info = Path.GetFileNameWithoutExtension(file).Replace("slide", string.Empty);
            var slideNr = int.Parse(info);
            return slideNr;
        }

        static void SaveSlidesAsImages(string pptPath, string outDir)
        {
            Application pptApplication = new Application();
            Presentation pptPresentation = pptApplication.Presentations.Open(pptPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            for (int i = 1; i <= pptPresentation.Slides.Count; i++)
            {
                var slide = pptPresentation.Slides[i];
                Console.WriteLine($"creating thumbnail {i} of {pptPresentation.Slides.Count-1}");
                string fileName = Path.Combine(outDir, $"slide{i}.png");
                string tmpFileName = fileName + ".tmp";
                slide.Export(tmpFileName, "png", 32, 38); // height 38px -> 32px: remove slide title area and dacadoo footer. title should be covered by text match already
                using (var img = Image.FromFile(tmpFileName))
                {
                    using (var bmpImage = new Bitmap(img))
                    {
                        using (var bmpCrop = bmpImage.Clone(new Rectangle(0, 4, 32, 30), bmpImage.PixelFormat))
                        {
                            bmpCrop.Save(fileName);
                        }
                    }
                }
                File.Delete(tmpFileName);
            }
        }

        private static IDictionary<int, string> ExtractText(string slides)
        {
            var lookup = new Dictionary<int, string>();

            using (ZipArchive zip = ZipFile.Open(slides, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    // "./ppt/slides/slide{n}.xml"
                    if (Path.GetExtension(entry.FullName).ToLowerInvariant() == ".xml" && entry.FullName.ToLowerInvariant().Contains(@"ppt/slides/slide"))
                    {
                        var stream = entry.Open();
                        var doc = new XmlDocument();
                        doc.Load(stream);
                        var info = Path.GetFileNameWithoutExtension(entry.Name).Replace("slide", string.Empty);
                        var slideNr = int.Parse(info);
                        lookup[slideNr] = GetText(doc);
                    }
                }
            }
            return lookup;
        }

        // slide nr -> text
        private static Dictionary<int, string> ReadFilesOld(string directory)
        {
            
            var lookup = new Dictionary<int, string>();
            var srcSlides = Directory.GetFiles(directory, "*.xml");
            foreach (var file in srcSlides)
            {
                var text = GetTextOld(file);
                var info = Path.GetFileNameWithoutExtension(file).Replace("slide", string.Empty);
                var slideNr = int.Parse(info);
                lookup[slideNr] = text;
            }
            return lookup;
        }

        private static string GetTextOld(string file)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(file);
            return GetText(doc);
        }

        private static string GetText(XmlDocument doc)
        {
            var text = new StringBuilder();
            var nodes = doc.GetElementsByTagName("t", "http://schemas.openxmlformats.org/drawingml/2006/main");
            if (nodes.Count > 0)
            {
                for (int i = 0; i < nodes.Count; i++)
                {
                    var node = nodes[i];
                    if (!string.IsNullOrWhiteSpace(node.InnerText))
                    {
                        text.Append(" ");
                        text.Append(node.InnerText);
                    }

                }
            }
            return text.ToString().Trim().ToLowerInvariant();
        }
    }
}
