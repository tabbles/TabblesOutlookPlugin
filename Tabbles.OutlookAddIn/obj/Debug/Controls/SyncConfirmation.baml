   M S B A M L   `   `   `  ����  TaggerOutlookAddIn4    .Tabbles.OutlookAddIn.Controls.SyncConfirmation   5   	   W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35\ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5      R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5      Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5      @d2http://schemas.microsoft.com/expression/blend/2008     5      .��5      �   � ColumnDefinitions  5	   
   �� 5
      $��50 �6   �� 5      $
��Auto �6   �� 5      $��50 �6    � RowDefinitions 5   
   �� 5      $��20 �6   �� 5      $
��Auto �6   �� 5      $��35 �6   �� 5      $
��Auto �6   �� 5      $��10 �6   �� 5      $
��Auto �6   �� 5      $��20 �6   .+�5   
   �� -   5         ��Name$ 
txbWarning��5      $��1��6(   $��1��68   $��Do you want to sync?��6E   � 5   
   -   5         �Name$ grdWarningButtons��5      $��1��6*   $��3��6:     5      �� 5      $��30 �6#   �� 5      $
��Auto �6#   �� 5      $��20 �6#   �� 5      $
��Auto �6#   �� 5       $��30 �6#   .+�5#      �� -   5         ��Name$ btnOk��5#      $��1��6$   $��60��64   $��OK��6?   6L   �� 5$      -   5        $ 	btnCancel��5$      $��3��6(   $��60��68   $��Cancel��6C    �� IsCancel . 6d   5%      �� 5'   
   -   5         ��Name$ chkDontAskAgain��5'      $��1��6,   $��5��6<   $��Don't ask again.��5(      6P   5)      5*      