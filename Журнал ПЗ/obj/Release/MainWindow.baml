   M S B A M L   `   `   `  ���� L  HЖурнал ПЗ, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null5clr-namespace:График_ПЗГрафик_ПЗ  R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R8clr-namespace:System.ComponentModel;assembly=WindowsBaseSystem.ComponentModel R NMoralesLarios.Utilities, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullnLclr-namespace:MoralesLarios.Utilities.Excel;assembly=MoralesLarios.UtilitiesMoralesLarios.Utilities.Excel "    График_ПЗ.MainWindow   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     @d2http://schemas.microsoft.com/expression/blend/2008     Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     1localclr-namespace:График_ПЗ     Hscm8clr-namespace:System.ComponentModel;assembly=WindowsBase     [mlLclr-namespace:MoralesLarios.Utilities.Excel;assembly=MoralesLarios.Utilities     -     � Icon$  Source/table_1.ico�� � Title$0 *Журнал проверки знаний�� � WindowStyle$ ThreeDBorderWindow=� � WindowState$ 	Maximized=�$��1466.5��$	��779�� � 	Resources    WrappedColumnHeaderStyle �� 
TargetType=  7System.Windows.Controls.Primitives.DataGridColumnHeader �� Property �� Value �� TextAlignment��%V   &	        �� "  .���  � �� �� .R��� $
��Wrap=�$ Center=����� .��� $��Stretch=�$��-2,0,2,0q�.+�o� $��24��$	��Top=�$
��Left=�$��55��.@�m� $��	Файл ��$��14(�.@�m� -   $��Загрузить��	 m� Icon	 �� $��Source/upload-button.png��m� $��Сохранить как��	 �� $ ��Source/direct-download.png��.@�m� -   $��.xlsx��	 �� $��Source/excel.png��m� -   $
��.pdf��	 �� $��Source/pdf.png��m� -   $��Печать��	 �� $��Source/printing.png���� -   $��57,1,0,0q�$
��Left=�$��24��$��24��$	��Top=�
�������.���� .+��� $��Source/plus.png���� -   $��84,1,0,0q�$
��Left=�$��24��$��24��$	��Top=�
�������.���� .+��� $��Source/delete.png���� -   $��	111,1,0,0q�$
��Left=�$��24��$��24��$	��Top=�
�������.���� .+��� $��Source/go-back-arrow.png���� $��
0,24,200,0q�.@��� $��0,0,0,0q����� -	   
 ��Name$
 label1��$��Center=�$��Все данные�����	#FFD97DE8$��90��$
��Bold"�.��&   System.Windows.Controls.DataGrid  -
     Name$ EmployeeDataGrid��   CanUserReorderColumns .  $��Stretch=�$	��Top=� J� RenderTransformOrigin$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)�$��0q� J� SnapsToDevicePixels .  � ContextMenu �� .@�m� -   $#��Вставить строку��	 �� $��Source/plus.png��m� -   $!��Удалить строку��	 �� $��Source/delete.png��m� -    m� Command$
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   Columns 0  *System.Windows.Controls.DataGridTextColumn  ,  &System.Windows.Controls.DataGridColumn   Header$ №п/п��1  +System.Windows.Controls.DataGridBoundColumn   Binding ��  �� Path$ OrderId� �� UpdateSourceTrigger$ PropertyChanged=�  $ ФИО�� �� $
 Name�$ PropertyChanged=�  -   
posCol��$, &Должность, профессия��   Width5  /System.Windows.Controls.DataGridLengthConverter$	 200  �� $ Position�$ PropertyChanged=�   ElementStyle �� .���  � ��$
 Wrap=���  � ��$ Center=�  $7 1Структурное подразделение�� �� $ 
Department�$ PropertyChanged=�     HeaderStyle# [     SortMemberPath$  ��$� �Дата проверки знаний по ОТ и оказанию первой помощи (после прохождения 40-часового обучения) факт��$	 100  �� $ ExaminationDateFact� �� StringFormat$ {0:dd.MM.yyyy}��$ PropertyChanged=� �� .���  � ��$
 Wrap=�  # [  $ 
dd.MM.yyyy��$� �Дата проверки знаний по ОТ и оказанию первой помощи (после прохождения 40-часового обучения) план��$	 100  �� $ ExaminationDatePlan�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   examComplFactCol��# [  $� �Дата проверки знаний по ОТ, ПУЭ, ПТЭ, ПожБ (согласно Правил работы с персоналом) факт��$	 100  �� $  ExaminationComplexDateFact�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   examComplPlanCol��# [  $� �Дата проверки знаний по ОТ, ПУЭ, ПТЭ, ПожБ (согласно Правил работы с персоналом) план��$	 100  �� $  ExaminationComplexDatePlan�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   
attestFact��# [  $c ]Дата аттестации по промышленной безопасности факт��$	 100  �� $ AttestationDateFact�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   
attestPlan��# [  $c ]Дата аттестации по промышленной безопасности план��$	 100  �� $ AttestationDatePlan�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   pbMinFactCol��# [  $g aДата прохождения пожарно-технического минимума факт��$	 100  �� $ PbminimumPassDateFact�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   pbMinPlanCol��# [  $g aДата прохождения пожарно-технического минимума план��$	 100  �� $ PbminimumPassDatePlan�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   
medPlanCol��# [  $	 100 $C =Дата проведения медосмотра предв�� �� $ MedicalCheckDateFact�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   
medFactCol��# [  $	 100 $E ?Дата проведения медосмотра период�� �� $ MedicalCheckDatePlan�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   
tabCol��# [  $ 75 $# Табельный номер�� �� $ 	TabNumber�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   birthCol��# [  $	 100 $ Дата рождения�� �� $ 	BirthDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   entryCol��$	 100 $K EДата проведения вводного инструктажа�� �� $ 	EntryDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   relocCol��# [  $	 100 $ Дата перевода�� �� $ RelocationDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  -   primCol��# [  $	 100 $O IДата проведения первичного инструктажа�� �� $ PrimaryInstructionDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ Стажировка��$	 100  �� $ InternshipDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $< 6Стажировка реквизиты приказа��# [  $	 100  �� $ InternshipDetails�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ Дублирование��$	 100  �� $ DublicationDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $@ :Дублирование реквизиты приказа��# [  $	 100  �� $ DublicationDetails�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ Допуск��$	 100  �� $ IndependentDate�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $4 .Допуск реквизиты приказа��# [  $	 100  �� $ IndependentDetails�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $= 7Расторгнут договор, переведен��# [  $	 100  �� $ ExtraStatus�$ PropertyChanged=� �� .���  � ��$
 Wrap=��� $��0,0,0,0q����� -   $
 tabitem2��$��Center=�$��ПЗ по ОТ и ПРП�����	#FF878DFF$	��107��$
��Bold"�.��  -   $ EmployeeDataGrid1�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -   $#��Вставить строку��	 �� $��Source/plus.png��m� -    $!��Удалить строку��	 �� $��Source/delete.png��m� -!   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -"   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� $ OrderId�$ PropertyChanged=�  $ ФИО�� �� $
 Name�$ PropertyChanged=�  -#   posCol2��$ Должность��$	 200  �� $ Position�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� $ 
Department�$ PropertyChanged=�  $ ПЗ по ОТ факт�� �� $ ExaminationDateFact�$ PropertyChanged=�  $ ПЗ по ОТ план�� �� $ ExaminationDatePlan�$ PropertyChanged=�  $ ПЗ по ПРП факт�� �� $  ExaminationComplexDateFact�$ PropertyChanged=�  $ ПЗ по ПРП план�� �� $  ExaminationComplexDatePlan�$ PropertyChanged=��� $��0,0,0,0q����� -$   $
 tabitem4��$��Center=�$&�� Аттестация по ПрБ�� � ToolTip$Q KАттестация по промышленной безопасности�����	#FFFFFF50$
��Bold"�$	��123��.��  -%   $ EmployeeDataGrid3�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -&   $#��Вставить строку��	 �� $��Source/plus.png��m� -'   $!��Удалить строку��	 �� $��Source/delete.png��m� -(   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -)   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� $ OrderId�$ PropertyChanged=�  $ ФИО�� �� $
 Name�$ PropertyChanged=�  $ Должность��$	 200  �� $ Position�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� $ 
Department�$ PropertyChanged=�  $6 0Дата аттестации по ПБ факт��# [  $	 100  �� $ AttestationDateFact�$ PropertyChanged=�  $6 0Дата аттестации по ПБ план��# [  $	 100  �� $ AttestationDatePlan�$ PropertyChanged=��� $��0,0,0,0q����� -*   $
 tabitem5��$��Center=�$��ПТМ��$: 4Пожарно-технический минимум�����	#FFFF5252$��58��$
��Bold"�.��  -+   $ EmployeeDataGrid4�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -,   $#��Вставить строку��	 �� $��Source/plus.png��m� --   $!��Удалить строку��	 �� $��Source/delete.png��m� -.   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -/   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� $ OrderId�$ PropertyChanged=�  $ ФИО�� �� $
 Name�$ PropertyChanged=�  $ Должность��$	 200  �� $ Position�$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� $ 
Department�$ PropertyChanged=�  $H BДата прохождения пож.-техн. мин. факт��# [  $	 100  �� $ PbminimumPassDateFact�$ PropertyChanged=�  $H BДата прохождения пож.-техн. мин. план��# [  $	 100  �� $ PbminimumPassDatePlan�$ PropertyChanged=��� $��0,0,0,0q����� -0   $
 tabitem6��$��Center=�$��т/н��$# Табельный номер�����	#FFC2E5FD$��69��$
��Bold"�.��  -1   $ EmployeeDataGrid5�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -2   $#��Вставить строку��	 �� $��Source/plus.png��m� -3   $!��Удалить строку��	 �� $��Source/delete.png��m� -4   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -5   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $# Табельный номер��# [  $	 100  �� *	TabNumber+$ PropertyChanged=��� $��0,0,0,0q����� -6   $
 	tabitem11��$��Center=�$)��#Вводный инструктаж�����	#FFC2E5FD$	��136��$
��Bold"�.��  -7   $ EmployeeDataGrid10�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -8   $#��Вставить строку��	 �� $��Source/plus.png��m� -9   $!��Удалить строку��	 �� $��Source/delete.png��m� -:   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -;   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $K EДата проведения вводного инструктажа��# [  $	 100  �� *	EntryDate+$ PropertyChanged=��� $��0,0,0,0q����� -<   $
 tabitem7��$��Center=�$��Медосмотр�����	#FFC2E5FD$��86��$
��Bold"�.��  -=   $ EmployeeDataGrid6�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� ->   $#��Вставить строку��	 �� $��Source/plus.png��m� -?   $!��Удалить строку��	 �� $��Source/delete.png��m� -@   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -A   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $C =Дата проведения медосмотра предв��# [  $	 100  �� *MedicalCheckDateFact+$ PropertyChanged=�  $E ?Дата проведения медосмотра период��# [  $	 100  �� *MedicalCheckDatePlan+$ PropertyChanged=��� $��0,0,0,0q����� -B   $
 tabitem8��$��Center=�$��Стажировка�����	#FFC2E5FD$��85��$
��Bold"�.��  -C   $ EmployeeDataGrid7�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -D   $#��Вставить строку��	 �� $��Source/plus.png��m� -E   $!��Удалить строку��	 �� $��Source/delete.png��m� -F   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -G   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $. (Стажировка даты с_ по _��# [  $	 150  �� *InternshipDate+$ PropertyChanged=�  $< 6Стажировка реквизиты приказа��# [  $	 150  �� *InternshipDetails+$ PropertyChanged=��� $��0,0,0,0q����� -H   $
 tabitem9��$��Center=�$��Дублирование�����	#FFC2E5FD$	��100��$
��Bold"�.��  -I   $ EmployeeDataGrid8�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -J   $#��Вставить строку��	 �� $��Source/plus.png��m� -K   $!��Удалить строку��	 �� $��Source/delete.png��m� -L   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -M   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $1 +Дублирование даты с_ по_��# [  $	 150  �� *DublicationDate+$ PropertyChanged=�  $@ :Дублирование реквизиты приказа��# [  $	 150  �� *DublicationDetails+$ PropertyChanged=��� $��0,0,0,0q����� -N   $
 	tabitem10��$��Center=�$��Допуск��$A ;Допуск к самостоятельной работе�����	#FFC2E5FD$��72��$
��Bold"�.��  -O   $ EmployeeDataGrid9�� .  $��Stretch=�$	��Top=�$ 	1.39,1.72&�$��0,0,0,0q�$��Times New Roman)� �� .@�m� -P   $#��Вставить строку��	 �� $��Source/plus.png��m� -Q   $!��Удалить строку��	 �� $��Source/delete.png��m� -R   $
 Copy��$'��!Копировать ячейку��	 �� $��Source/minus.png��m� -S   $
 Copy��$'��!Копировать строку��	 �� $��Source/cells.png��   $ №п/п�� �� *	OrderId+$ PropertyChanged=�  $ ФИО�� �� *Name+$ PropertyChanged=�  $ Должность��$	 200  �� *
Position+$ PropertyChanged=� �� .���  � ��$
 Wrap=�  $ 
Отдел�� �� *
Department+$ PropertyChanged=�  $ Допуск дата��# [  $	 150  �� *IndependentDate+$ PropertyChanged=�  $4 .Допуск реквизиты приказа��# [  $	 150  �� *IndependentDetails+$ PropertyChanged=��� $��3��$��0,55,0,0q�$	��Top=�$��Right=�$	��200���� -T   $��	0,79,17,0q�$	��Top=�
�������
���   �$��41��$��14(����	#FFFFBE71$��Right=�$	��167��.���� $8��2Истек срок проверки знаний
��$ Center=�$
��Wrap=�$��Medium"��� $��4��$��	0,141,0,0q�$	��Top=�$��Right=�$	��200���� -U    ��Name$ searchString��$��24��$��	213,0,0,0q� �� TextWrapping$
 Wrap=�$	��Top=�$. (Введите строку поиска��$
��Left=�$	��156��(  "System.Windows.Controls.DatePicker  -V     Name$ to_date��$��Right=�$��	0,276,0,0q�$	��Top=�$	��111��$��29���� $��От:��$��Center=�$��Right=�$��0,240,111,0q�$	��Top=�$��89�����	#FFF4EDFD$��14(�$
��Bold"��� $��До:��$��Center=�$��Right=�$��0,277,111,0q�$	��Top=�$��89�����	#FFF4EDFD$��14(�$
��Bold"��� $��Center=�$��	0,150,0,0q�$	��Top=�$��69�����	#FFAFF99E$��Medium"�$��Right=�$	��200��.���� $S��MГрафик проверки знаний в выбранный период��$ Center=�$
��Wrap=�$��14(�$��55��$ 0.494,0.873&��� -W   $��Сформировать��$��
0,324,33,0q�$	��Top=�$��Right=�$	��122��
�������
���   �$��27��$��14(�  -X   $ 	from_date��$��Right=�$��	0,239,0,0q�$	��Top=�$	��111��$��29���� $��	180,0,0,0q�$��Source\search.png��$��24��$	��Top=�$ -1.423,0.403&�$
��Left=�$��26���� $��5��$��	0,368,0,0q�$	��Top=�$��Right=�$	��200���� -Y   ���	#FFC2E5FD$��Right=�$	��Top=�$��22��$��0,378,173,0q�$��20���� -Z   ���	#FFAFF99E$	��Top=�$��0,378,146,0q�$��20��$��Right=�$��22���� -[   ���#e9ff59$	��Top=�$��0,378,119,0q�$��20��$��Right=�$��22���� -\   ���	#FFFFBE71$	��Top=�$��
0,378,92,0q�$��20��$��Right=�$��22���� -]   
�������$	��Top=�$��
0,378,65,0q�$��20��$��Right=�$��22��