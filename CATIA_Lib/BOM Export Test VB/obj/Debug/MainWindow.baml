   M S B A M L   `   `   `  ����M  IBOM Export Test VB, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null7 clr-namespace:BOM_Export_Test_VBBOM_Export_Test_VB  5   	   #    BOM_Export_Test_VB.MainWindow   5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   2local clr-namespace:BOM_Export_Test_VB     5   	     � Title$  
MainWindow��5   	   $	��350��6   $	��525��6)    � SizeToContent$ WidthAndHeight=�65   .��5	      � .+�5
   
   �� -   5         ��Name$ button��5
      $��Button��6!   $
��Left=�62   $��42��6M   $��366,101,0,0q�6Y   $	��Top=�6n   $��81��6�   �� 5   
   -   5         ��Name$ 
PartNumber��5      $
��Left=�6(   $��34��6C   $��
18,109,0,0q�6O   $
��Wrap=�6c   $��	TextBlock��6w   $	��Top=�6�   $	��219��6�   �� 5   
   -   5        $ button1��5      $��3D��6"   $
��Left=�6/   $��54��6J   $��
180,89,0,0q�6V   $	��Top=�6j   $��95��6�   &   System.Windows.Controls.DataGrid  5   
   -   5          Name$ dataGrid��5      $
��Left=�6%   $
��Auto��6@   $��
18,186,0,0q�6N   $	��Top=�6b   $
��Auto��6z   �� 5   
   -   5        $ button2��5      $��Create Panel��6"   $
��Left=�69   $��55��6T   $��237,216,0,0q�6`   $	��Top=�6u   $��91��6�   �� 5   
   -   5        $ button3��5      $��New Product��6"   $
��Left=�68   $��34��6S   $��
111,22,0,0q�6_   $	��Top=�6s   $	��164��6�   �� 5   
   -   5        $ button4��5      $��Clash��6"   $
��Left=�62   $��56��6M   $��389,199,0,0q�6Y   $	��Top=�6n   $	��105��6�   �� 5   
   -   5        $ button5��5      $��Write Notes��6"   $
��Left=�68   $��34��6S   $��
351,22,0,0q�6_   $	��Top=�6s   $	��119��6�   5      5      