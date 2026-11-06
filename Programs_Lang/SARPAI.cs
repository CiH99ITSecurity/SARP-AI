using System;
using System.IO;
using System.Drawing;

using Xceed.Words.NET;
using Xceed.Document.NET;

namespace SARPAI
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // تحديد مجلد العمل
                string basePath = AppDomain.CurrentDomain.BaseDirectory;

                // مسار قالب Word
                string templatePath = Path.Combine(
                    basePath, "Template.docx");

                // مسار الصور
                string image1Path = Path.Combine(
                    basePath, "Placeholders",
                    "Placeholder_1.png");

                string image2Path = Path.Combine(
                    basePath, "Placeholders",
                    "Placeholder_2.png");

                // مسار الملف الناتج
                string outputPath = Path.Combine(
                    basePath, "Template_Rendered.docx");

                // التحقق من وجود الملفات
                if (!File.Exists(templatePath))
                    throw new FileNotFoundException(templatePath);

                if (!File.Exists(image1Path))
                    throw new FileNotFoundException(image1Path);

                if (!File.Exists(image2Path))
                    throw new FileNotFoundException(image2Path);

                // فتح قالب Word
                using (DocX doc = DocX.Load(templatePath))
                {
                    // استبدال الاسم
                    doc.ReplaceText(
                        "{{ name }}",
                        "SARPAI"
                    );

                    // إدراج الصورة الأولى
                    InsertImage(
                        doc,
                        "{{ placeholder_1 }}",
                        image1Path,
                        5.0f
                    );

                    // إدراج الصورة الثانية
                    InsertImage(
                        doc,
                        "{{ placeholder_2 }}",
                        image2Path,
                        5.0f
                    );

                    // حفظ المستند
                    doc.SaveAs(outputPath);
                }

                Console.WriteLine(
                    "Document created successfully!");

                Console.WriteLine(outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(
                    "Error: " + ex.Message);
            }

            Console.ReadKey();
        }

        // إدراج صورة داخل قالب Word
        private static void InsertImage(
            DocX doc,
            string placeholder,
            string imagePath,
            float widthCm)
        {
            // حساب أبعاد الصورة
            float widthPoints =
                widthCm * 72f / 2.54f;

            float heightPoints;

            using (Image img = Image.FromFile(imagePath))
            {
                heightPoints =
                    widthPoints * img.Height / img.Width;
            }

            // تحميل الصورة
            var image = doc.AddImage(imagePath);

            // إنشاء الصورة بالحجم المطلوب
            Picture picture = image.CreatePicture(
                widthPoints,
                heightPoints
            );

            // استبدال العلامة بالصورة
            doc.ReplaceTextWithObject(
                new ObjectReplaceTextOptions()
                {
                    SearchValue = placeholder,
                    NewObject = picture
                }
            );
        }
    }
}