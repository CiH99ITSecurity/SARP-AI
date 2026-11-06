Imports System
Imports System.IO
Imports System.Drawing

Imports Xceed.Words.NET
Imports Xceed.Document.NET

Module Program

    Sub Main()

        Try
            ' تحديد مجلد العمل
            Dim basePath As String =
                AppDomain.CurrentDomain.BaseDirectory

            ' مسارات الملفات
            Dim templatePath As String =
                Path.Combine(basePath, "Template.docx")

            Dim image1Path As String =
                Path.Combine(basePath, "Placeholders",
                             "Placeholder_1.png")

            Dim image2Path As String =
                Path.Combine(basePath, "Placeholders",
                             "Placeholder_2.png")

            Dim outputPath As String =
                Path.Combine(basePath, "Template_Rendered.docx")

            ' التحقق من وجود الملفات
            If Not File.Exists(templatePath) Then
                Throw New FileNotFoundException(templatePath)
            End If

            If Not File.Exists(image1Path) Then
                Throw New FileNotFoundException(image1Path)
            End If

            If Not File.Exists(image2Path) Then
                Throw New FileNotFoundException(image2Path)
            End If

            ' فتح قالب Word
            Using doc As DocX = DocX.Load(templatePath)

                ' استبدال الاسم
                doc.ReplaceText(
                    "{{ name }}",
                    "SARPAI"
                )

                ' إدراج الصورة الأولى
                InsertImage(
                    doc,
                    "{{ placeholder_1 }}",
                    image1Path,
                    5.0F
                )

                ' إدراج الصورة الثانية
                InsertImage(
                    doc,
                    "{{ placeholder_2 }}",
                    image2Path,
                    5.0F
                )

                ' حفظ المستند
                doc.SaveAs(outputPath)

            End Using

            Console.WriteLine(
                "Document created successfully!"
            )

            Console.WriteLine(outputPath)

        Catch ex As Exception

            Console.WriteLine(
                "Error: " & ex.Message
            )

        End Try

        Console.ReadKey()

    End Sub


    ' إدراج صورة داخل مستند Word
    Private Sub InsertImage(
        ByVal doc As DocX,
        ByVal placeholder As String,
        ByVal imagePath As String,
        ByVal widthCm As Single
    )

        ' تحويل السنتيمتر إلى Points
        Dim widthPoints As Single =
            widthCm * 72.0F / 2.54F

        Dim heightPoints As Single

        ' حساب ارتفاع الصورة مع المحافظة على النسبة
        Using img As System.Drawing.Image =
            System.Drawing.Image.FromFile(imagePath)

            heightPoints =
                widthPoints * CSng(img.Height) /
                CSng(img.Width)

        End Using

        ' تحميل الصورة
        Dim image = doc.AddImage(imagePath)

        ' تحديد أبعاد الصورة
        Dim picture As Picture =
            image.CreatePicture(
                widthPoints,
                heightPoints
            )

        ' استبدال العلامة بالصورة
        Dim options As New ObjectReplaceTextOptions With {
            .SearchValue = placeholder,
            .NewObject = picture
        }

        Dim replaced As Boolean =
            doc.ReplaceTextWithObject(options)

        If Not replaced Then
            Throw New Exception(
                "Placeholder not found: " & placeholder
            )
        End If

    End Sub

End Module