# 教程：使用Python向Docx文档嵌入附件

## 目的


## 一、准备
office，python，olefile，oletools

## 二、认识docx文档
docx格式的word文档其实就是使用zip压缩的以xml文档为主的文件目录。
1. 首先，我们来创建一个word文档，随便输入几个字符，保存为“demo.docx”。
2. 然后，我们使用python来解压该文件。当然，你也可以使用解压工具来完成这个工作。
   - 创建“word”文件夹，把<code>demo.docx</code>和下面的python脚本都放到该文件夹下。
   - 编写<code>embeddocx-01.py</code>脚本：

        ``` python
        # embeddocx-01.py
        import os
        import shutil
        import zipfile

        docx_fn = 'demo.docx'
        extract_folder = 'extrated'

        def unzip_docx():
            shutil.rmtree(extract_folder, ignore_errors=True)
            os.mkdir(extract_folder)
            os.chdir(extract_folder)
            fn = os.path.join('../', docx_fn)
            with zipfile.ZipFile(fn) as azip:
                azip.extractall()


        if __name__ == '__main__':
            unzip_docx()
        ```
    - 到word文件夹下运行该脚本
        ``` bash
        cd word
        python embeddocx-01.py
        ```

3. 该脚本会将docx文件解压到extracted文件夹下，我们查看下extracted的目录结构。其中这三个文件和嵌入附件相关：<code>[Content_Types].xml</code>、<code>document.xml.rels</code>及<code>document.xml</code>。我们把它们拷贝出来以便和后面的对比。
   
   <image src="01.png" width="800">

## 三、嵌入文件
1. 向demo.docx文件添加一个word文件，选择“显示为图标”，然后再次保存。
1. 再次运行python脚本，然后再次观察extracted目录结构，此时会发现word子文件夹下多了embeddings和media两个子文件夹。
   
   <image src="02.png" width="240">

1. 试着打开embeddings下的Microsoft_Word___.docx，你会发现正是我们刚才嵌入的文档。而media下的image1.emf则是在word中显示的图标，是一个矢量格式的图形文件。
7. 比较<code>[Content_Types].xml</code>文件，我们会发现多了2行：
   ```xml
    <Default Extension="emf" ContentType="image/x-emf"/>
    <Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>
   ```
   
8. 比较<code>document.xml.rels</code>文件，我们会发现也多了2行。注意其中的“Id”的值：
   ```xml
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/Microsoft_Word___.docx"/>
    <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>
   ```
   
9.  比较<code>document.xml</code>文件，我们会发现多了如下部分。注意其中的“ProgID”、“r:Id”的值：
   ```xml
    <w:object w:dxaOrig="1485" w:dyaOrig="1005">
        <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
            <v:stroke joinstyle="miter"/>
            <v:formulas>
                <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                <v:f eqn="sum @0 1 0"/>
                <v:f eqn="sum 0 0 @1"/>
                <v:f eqn="prod @2 1 2"/>
                <v:f eqn="prod @3 21600 pixelWidth"/>
                <v:f eqn="prod @3 21600 pixelHeight"/>
                <v:f eqn="sum @0 0 1"/>
                <v:f eqn="prod @6 1 2"/>
                <v:f eqn="prod @7 21600 pixelWidth"/>
                <v:f eqn="sum @8 21600 0"/>
                <v:f eqn="prod @7 21600 pixelHeight"/>
                <v:f eqn="sum @10 21600 0"/>
            </v:formulas>
            <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
            <o:lock v:ext="edit" aspectratio="t"/>
        </v:shapetype>
        <v:shape id="_x0000_i1031" type="#_x0000_t75" style="width:74.25pt;height:50.25pt" o:ole="">
            <v:imagedata r:id="rId6" o:title=""/>
        </v:shape>
        <o:OLEObject Type="Embed" ProgID="Word.Document.12" ShapeID="_x0000_i1031" DrawAspect="Icon" ObjectID="_1673767165" r:id="rId7">
            <o:FieldCodes>\s</o:FieldCodes>
        </o:OLEObject>
    </w:object>
   ```

1. 这时候，可能你们已经看出来了，嵌入一个word文件到docx文件中，会有以下变化：
   1. word文件本身会放到embeddings子文件夹下
   2. 图标文件会放到media子文件夹下
   3. <code>[Content_Types].xml</code>文件中多了2行定义，是为了支持上面的附件和图标文件格式
   4. <code>document.xml.rels</code>文件中也多了2行，为上面2个文件各分配了唯一的id，并指明了文件类型和路径
   5. <code>document.xml</code>文件中多了“w:object”部分，其中有：
      - “v:imagedata r:id="rId6"”指出了图标文件的引用的id，该id在<code>document.xml.rels</code>中定义了其文件类型和文件路径
      - “o:OLEObject r:id="rid7"”指出了嵌入文件的引用的id，该id在<code>document.xml.rels</code>中定义了其文件类型和文件路径

2. 有些复杂，参数也很多。到底哪些内容和参数对我们有影响呢？让我们做些改动。
   1. 将embeddings下的“Microsoft_Word___.docx”文件名修改为“file2001.docx”，相应地<code>document.xml.rels</code>里面的Target也跟着修改。
   2. 将下面的图片保存名为“doc.png”，并拷贝到media下。相应地<code>document.xml.rels</code>里面的Target也跟着修改。

        <image src="doc.png" width="40">

      修改后的部分如下：
        ```xml
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="embeddings/file2001.docx"/>
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/doc.png"/>
        ```

   3. 修改<code>[Content_Types].xml</code>文件，添加对PNG图片的支持：
      ```xml
       <Default Extension="png" ContentType="image/png"/>
      ```
   4. 修改<code>document.xml</code>文件中“w:object”部分内容：
      - 删除“v:shapetype”部分
      - 将“v:shape”的id和“o:OLEObject”的ShapeID都修改为"_x0000_i2001"
      - 将“v:shape”的style的width和height都修改为48
      - 将“o:OLEObject”的ObjectID修改为"_1673762001"     
        ```xml
        <w:object w:dxaOrig="1485" w:dyaOrig="1005">
            <v:shape id="_x0000_i2001" type="#_x0000_t75" style="width:74.25pt;height:50.25pt" o:ole="">
                <v:imagedata r:id="rId6" o:title=""/>
            </v:shape>
            <o:OLEObject Type="Embed" ProgID="Word.Document.12" ShapeID="_x0000_i2001" DrawAspect="Icon" ObjectID="_1673762001" r:id="rId7">
            </o:OLEObject>
        </w:object>
        ```

3. 然后，编写并运行python脚本，将修改后的文件重新压缩成docx文件。

    ```python
    # embeddocx-02.py
    import os
    import shutil
    import zipfile

    docx_fn = 'demo.docx'
    extract_folder = 'extrated'
    this_path = os.path.dirname(os.path.abspath(__file__))
    src_docx_fn = os.path.join(this_path, docx_fn)


    def unzip_docx():
        shutil.rmtree(extract_folder, ignore_errors=True)
        os.mkdir(extract_folder)
                                os.chdir(extract_folder)
        with zipfile.ZipFile(src_docx_fn) as azip:
            azip.extractall()


    def zip_docx():
        new_docx_fn = os.path.join(this_path, 'demo1.docx')
        os.chdir(extract_folder)
        with zipfile.ZipFile(new_docx_fn, 'w') as azip:
            for i in os.walk('.'):
                for j in i[2]:
                        azip.write(os.path.join(i[0], j),
                                compress_type=zipfile.ZIP_DEFLATED)


    if __name__ == '__main__':
        zip_docx()
    ```

4. 新的“demo1.docx”会生成，使用word打开，你会发现图标换了，双击图标，内嵌的word文件能够正常打开。

5. 小结：
   - docx文档就是使用zip压缩的带目录结构的一堆文件集合。
   - 嵌入的附件放在embeddings目录下，文件名可以自己指定（注意不要使用中文字符）。
   - 图标放在media目录下，可以自己指定图标。
   - <code>[Content_Types].xml</code>、<code>document.xml.rels</code>及<code>document.xml</code>这三个文件和嵌入附件相关。
   - 包括图标在内的所有文件类型必须在<code>[Content_Types].xml</code>定义扩展，常用的类型参考如下：
        ```xml
        <Default Extension="png" ContentType="image/png"/>
        <Default Extension="jpg" ContentType="image/jpeg"/>
        <Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Default Extension="docx" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"/>
        <Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>
        <Default Extension="xlsm" ContentType="application/vnd.ms-excel.sheet.macroEnabled.12"/>
        <Default Extension="pptx" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation"/>
        <Default Extension="doc" ContentType="application/msword"/>
        <Default Extension="xls" ContentType="application/vnd.ms-excel"/>
        <Default Extension="ppt" ContentType="application/vnd.ms-powerpoint"/>
        ```
        想要获取其它类型，可以将文档嵌入到docx文件中，保存再解压。然后打开<code>[Content_Types].xml</code>文件查看。
   - <code>document.xml.rels</code>文件定义了嵌入文档和图标的文件类型及路径，以及引用的id。想要获取其它格式文件的类型，可以将文档嵌入到docx文件中，保存后再解压，然后打开<code>document.xml.rels</code>文件查看。
   - <code>document.xml</code>文件的“w:object”中通过r:id来对应<code>document.xml.rels</code>中指定的文档和图标。
   - <code>document.xml</code>文件的“w:object”中有些参数值并不会影响嵌入的文档，只要其和其它不冲突就可以。

## 四、小试身手：插入Office文档
参考上面步骤，自己尝试嵌入一个Excel或其它office文件。你可以修改图标、文件名，甚至是引用的Id。

## 五、插入其它格式文件
参考上面步骤，添加一个zip文件。

### 初识Ole

### 解决方案


## 遗留的问题


## 总结

## 参考资料
