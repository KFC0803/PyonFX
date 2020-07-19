.. _quick-start:

快速入门指南
-----------------

首先，您需要知道您要如何进行创作。您需要了解（如果您还不知道的话）如下内容：

* **ASS格式** Pyonfx仍然是一个高级的排版工具和卡拉OK特效（karaokers）工具，它应该被经验丰富的用户（typesetters）使用，他们知道Libass所提供的所有标签。检查脚注以获得所有标签的完整列表。
* **Python3 语言** 像Python这样的编程语言可以让您定义在某种情况下您想要什么、多久一次、应用于这个或那个……总得来说能使您更自由。在一个完全图形化的界面中，您将不局限于只有几个命令的按钮、滑块或文本字段。**基本上就是这样**。变量、函数、条件、循环、比较、字符串格式化、列表和字典……您可以在脚注中找到一些优秀教程的链接。 [#f2]_

要开始生成，您只需像以前那样在Python 3中编写脚本，它包含您的KFX或高级排版创作的一切。

如果您还不会安装Python，您可以从网上找点教程，比如 https://www.liaoxuefeng.com/wiki/1016959663602400/1016959856222624。

Windows
+++++++

总之，如果您还没 **安装Python3** ， 
您可以从 `官网 <https://www.python.org/downloads/>`_ 上 **下载** 。
请确保勾选“Add Python 3.x to PATH”复选框。这使得您不需要进行额外操作来使Python可以在命令行中被调用。


运行如下命令，该命令将使用 pip 安装并最终更新库：

.. code-block:: sh
   :emphasize-lines: 1

   pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

好了。一切准备就绪，当您需要更新时，就再运行一次上面的命令。

Ubuntu/Debian
+++++++++++++

警告：下面的第一个命令还未完全测试。如果您遇到任何问题，请创建一个issue或参考 `官方的安装指南 <https://pygobject.readthedocs.io/en/latest/getting_started.html>`_ 。

.. code-block:: sh
   :emphasize-lines: 1,2
   
   sudo apt install python3 python3-pip libgirepository1.0-dev gcc libcairo2-dev pkg-config python3-dev gir1.2-gtk-3.0 python3-gi python3-gi-cairo
   python3 -m pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

Fedora
++++++

警告：下面的第一个命令还未完全测试。如果您遇到任何问题，请创建一个issue或参考 `官方的安装指南 <https://pygobject.readthedocs.io/en/latest/getting_started.html>`_ 。

.. code-block:: sh
   :emphasize-lines: 1,2
   
   sudo dnf install python3 python3-pip gcc gobject-introspection-devel cairo-devel pkg-config python3-devel python3-gobject gtk3
   python3 -m pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

Arch Linux
++++++++++

警告：下面的第一个命令还未完全测试。如果您遇到任何问题，请创建一个issue或参考 `官方的安装指南 <https://pygobject.readthedocs.io/en/latest/getting_started.html>`_ 。

.. code-block:: sh
   :emphasize-lines: 1,2
   
   sudo pacman -S python python-pip cairo pkgconf gobject-introspection python-gobject gtk3
   python3 -m pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

openSUSE
++++++++

警告：下面的第一个命令还未完全测试。如果您遇到任何问题，请创建一个issue或参考 `官方的安装指南 <https://pygobject.readthedocs.io/en/latest/getting_started.html>`_ 。

.. code-block:: sh
   :emphasize-lines: 1,2
   
   sudo zypper install python3 python3-pip cairo-devel pkg-config python3-devel gcc gobject-introspection-devel python3-gobject python3-gobject-Gdk typelib-1_0-Gtk-3_0 libgtk-3-0
   python3 -m pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

macOS
+++++

您可能需要先安装 `Homebrew <https://brew.sh/>`_ 。

警告：下面的第一个命令还未完全测试。如果您遇到任何问题，请创建一个issue或参考 `官方的安装指南 <https://pygobject.readthedocs.io/en/latest/getting_started.html>`_ 。

.. code-block:: sh
   :emphasize-lines: 1,2
   
   brew install pygobject3 gtk+3 cairo py3cairo pkg-config
   python3 -m pip install --upgrade https://github.com/CoffeeStraw/PyonFX/zipball/master

警告：如果输出没有正确渲染，可能需要将 PangoCairo 后端更改为 fontconfig。

.. code-block:: sh
   :emphasize-lines: 1
   
   PANGOCAIRO_BACKEND=fc python3 namefile.py


安装 - 额外步骤
+++++++++++++++++++++++++

跳过这一步并不影响此库的正常使用。但就我个人而言，我认为Aegisub太老/太臃肿了，我需要一种更舒适的工作方式。

这就是为什么PyonFX集成了一种额外的方式，您可以更快地在修改软字幕之后使用 `MPV播放器 <https://mpv.io/>`_ 。预览您的作品。如果您 *不* 使用Windows，安装完就可以了。

如果您正使用Windows，安装后（去网站上找），您需要把它添加到 PATH 中，这样库就可以使用它了。您可以在 `这里 <https://jingyan.baidu.com/article/8ebacdf02d3c2949f65cd5d0.html>`_ 找到一些指导。

您需要添加包含mpv的.exe的文件夹到PATH中，通常是C:\\Program Files\\mpv。


开始
++++++++

您可能想检查一下现在一切是否正常。为此，我建议您尝试运行一些在 `项目github官方仓库 <https://github.com/CoffeeStraw/PyonFX/tree/master/examples>`_ 中的示例。

要在Python中运行脚本，只需运行以下命令：

.. code-block:: sh
   :emphasize-lines: 1

   python namefile.py

或者，如果由于某种原因（例如您不使用Windows，或者同时装了Python2及Python3）而无法运行:

.. code-block:: sh
   :emphasize-lines: 1

   python3 namefile.py

我强烈建议您生成并研究examples文件夹中的每个示例（从 `'这里' <https://minhaskamal.github.io/DownGit/#/home?url=https://github.com/CoffeeStraw/PyonFX/tree/master/examples>`_ 下载最新的）。这些都是为零基础初学者进阶为高级用户准备的，并详细解释了库中所有相关函数的用法。

Tips
++++

* 不要一口气做一个KFX。休息一下，散散步，从周围环境中寻求灵感；
* 选择视频元素。您的效果应以某种方式与背景契合；
* 考虑人类的认知。我们主要关注运动，其次是对比度，然后是颜色。效果太多会让人头疼，少了又很无趣；
* 使用现代风格来升华（光，曲线，粒子，渐变），用传统风格来增强可读性（纯色，粗边框，静态位置）；
* 如果背景太华丽，请尝试插入面板形状 *(panel shape)* ，以将文本置于“安全地带”中；
* 适应卡拉OK时间和歌声。快歌没有音节持续时间，不适合使用需要一些时间才能看到的效果。

----------

.. rubric:: Footnotes
.. [#f1] ASS标签大全：https://aegi.vmoe.info/docs/3.2/ASS_Tags/
.. [#f2] 推荐的Python3 教程：https://www.liaoxuefeng.com/wiki/1016959663602400
