<%
'1.���SESSION USERNAMEֵΪ������ת����ҳ��
if session("UserName")=""  then                                                    '����û���SESSIONֵΪ��
response.write "<script>parent.location.href='pro/action.login.asp';</script>"     '��ת����ҳ��
response.End()                                                                     '�����ֹ
end if                                                                             '��������
'2.����SESSION����ʱ��
Session.Timeout=15                                                                 'SEESION��Чʱ��Ϊ15���� 
'3.��վ�����ļ�
'��վ����
WebTitle="���ۿ�����"


%>