�U���ѥ����:
�Ѽ�1:Stock.exe�ɮצ�m 
�Ѽ�2:Download
�Ѽ�3:�d�ߦ~(�̦�94/09/01�}�l)
�Ѽ�4:�d�ߤ�
�Ѽ�5:�d�ߤ�
�Ѽ�6:�������إN��
�Ѽ�7:��X�ɮצ�m(�̫�n��\)
�d��(106/04/07):Exe\Stock.exe Download 106 04 07 ALL D:\StockOutputData\Source\
�d��(105/10-105/12):
for /L %%m in (10 1 12) do (
	for /L %%d in (1 1 31) do (
		if %%m LSS 10 (
			if %%d LSS 10 (
				Exe\Stock.exe Download 105 0%%m 0%%d ALL D:\StockOutputData\Source\
			) else (
				Exe\Stock.exe Download 105 0%%m %%d ALL D:\StockOutputData\Source\
			)
		) else (
			if %%d LSS 10 (
				Exe\Stock.exe Download 105 %%m 0%%d ALL D:\StockOutputData\Source\
			) else (
				Exe\Stock.exe Download 105 %%m %%d ALL D:\StockOutputData\Source\
			)
		)
	)
)

���R�ѥ����:
�Ѽ�1:Stock.exe�ɮצ�m 
�Ѽ�2:Analyze
�Ѽ�3:�ӷ���Ƨ�
�Ѽ�4:�ؼи�Ƨ�(�̫�n��\)
�d��:Exe\Stock.exe Analyze D:\StockOutputData\Source D:\StockOutputData\

�`�N�ƶ�:
1.�u�B�z�ӷ���Ƨ������ɮ�(���]�t�l�h�ɮ�)
2.�����g�J��ؼи�Ƨ�����Ƴ��O���s�p���мg���A���|�֭p
3.Backup��Ƨ�����AnalysisData(��l���).xls���̪쪺�Ÿ�ơA�ƥ���
4.�C�����槹�ѥ����R��A�O�o�ƥ���Backup��Ƨ����A�Y���ᦳ���D�~���ΦA��������
5.��ƹJ��"-"��"0.00"�Ҥ��|�C�J�p��A#NUM!��ܰ��H0�����G
6.Exe��Ƨ��񪺬O�{����������