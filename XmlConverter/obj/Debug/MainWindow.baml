   M S B A M L   `   `   `  ����G  CXmlConverter, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null+clr-namespace:XmlConverterXmlConverter  5   	       XmlConverter.MainWindow   5   	   R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   ,localclr-namespace:XmlConverter     5   	     � Title$  XML Converter��5   	   $	��145��6   $	��718��6,   .��5	      � $��0,0,0,-3q�6    � ColumnDefinitions 5
   
   �� 5      $��0* �6   �� 5      .+�5   
   �� -   5        $��ConvertButton��5   '   $��Convert XML��6   $
��Left=�6O   $��24��6j   $��
591,50,0,0q�6v   $	��Top=�6�   $��99��6�   $��2��6�   �� 5   
   -   5        $��FilePath��5   -   $
��Left=�6   $��24��6=   $��	20,16,0,0q�6I    �� TextWrapping$
 Wrap=�6\   $��Path to XML File��6p   $	��Top=�6�   $	��554��6�   $��2��6�   �� 5   
   -   5        $��Status��5   /   $
��Left=�6   $��24��6=   $��	20,83,0,0q�6I   $
��Wrap=�6\   $�� ��6p   $	��Top=�6x   $	��554��6�   $��2��6�   	� 5   
   -   5        $	��PDF��5      $��
Xml to PDF��6!   $
��Left=�66   $��	20,55,0,0q�6Q   $	��Top=�6d   $��2��6|   	� 5   
   -   5        $
��Word��5      $��Xml to Word��6"   $
��Left=�68   $��
124,55,0,0q�6S   $	��Top=�6g   $��2��6    J� RenderTransformOrigin$ 0.068,-0.587&�6�   �� 5   
   -   5        $��Select XML File��5      $
��Left=�6>   $��
591,16,0,0q�6Y   $	��Top=�6m   $��99��6�   $��2��6�   $ 
0.475,0.05&�6�   $��24��6�   5      5      