<script Language=javaScript>
function ChkID(pUserID,pType) //�s�f(�����Ҧr��,�Τ@�s��)
{  
/*	�ǤJ�Ѽ�:	1.pUserID:�����Ҧr��:10���Ʀr,�Τ@�s��:8��Ʀr
				2.pType:0:�s�f�����Ҧr��,1:�s�f�Τ@�s��
	�Ǧ^�Ѽ�:Boolean:True�����T,False���T	*/
    var ix_I;
    if (pUserID == "") {return false;}
    switch(pType)
    {
		case 0: //�s�f�����Ҧr��
			var tAreaNo;
			var tCheckSum;
			var tAreaCode;
			var tSecondID;         //�����ҲĤG�X

			pUserID = pUserID.toUpperCase();
			tAreaCode = pUserID.substr(0,1);
			if (pUserID.length != 10)  //�T�w�����Ҧr����10�X
			{	alert("��J�L�Ī������Ҧr�� (ex:��ƪ��׿��~) !");
				return true;
			}
			if (tAreaCode.valueOf()<"A" && tAreaCode.valueOf()>"Z")  //�T�w���X�bA-Z����
			{	alert("��J�L�Ī������Ҧr�� (ex:���X������A-Z����) !");
				return true;
			}
			if (isNaN(parseInt(pUserID.substring(1,10),10)) == true)  //�T�w2-10�X�O�Ʀr
			{	alert("��J�L�Ī������Ҧr�� (ex:��2-10�X���O�Ʀr) !");
				return true;
			}
			//�����Ҹ��X�� 2 �X������ 1 �� 2
			tSecondID = pUserID.substr(1,1);
			if (tSecondID != 1 && tSecondID != 2) 
			{    alert("��J�L�Ī������Ҧr�� !");
			    return true;
			}
			//���o���X�������ϰ�X�AA ->10, B->11, ..H->17,I->34, J->18...
			tAreaNo = "ABCDEFGHJKLMNPQRSTUVXYWZIO".search(tAreaCode) + 10;
			pUserID = tAreaNo.toString(10) + pUserID.substring(1,10);   

			//  ���oCheckSum����,�ֹ鶴���Ҹ��X�O�_���T
			//  A = ��1�X, A0 = ��1�X*(10-1), A1 = ��2�X*(10-2), A2 = ��3�X*(10-3)
			//  A3 = ��4�X*(10-4), A4 = ��5�X*(10-5), A5 = ��6�X*(10-6)
			//  A6 = ��7�X*(10-7), A7 = ��8�X*(10-8), A8 = ��9�X*(10-9)
			//  CheckSum = A+A0+A1+A2+A3+A4+A5+A6+A7+A8

			tCheckSum = parseInt(pUserID.substr(0,1),10) + parseInt(pUserID.substr(10,1),10);
			for(ixI=1;ixI<=9;ixI++)
			{	tCheckSum = tCheckSum + parseInt(pUserID.substr(ixI,1),10)*(10-ixI);}
			if ((tCheckSum % 10) != 0)
			{    alert("��J�L�Ī������Ҧr�� !");
			     return true;
			}
			return false;
			break;

		case 1: //�s�f�Τ@�s��
			var tSum=0;
			var tDiv=0;
			var tMod=0;
			var tStr="12121241";
			         
			if (parseInt(pUserID.substring(0,8),10)!=pUserID) //�T�w1-8�X�O�Ʀr 
			{   alert("��J�L�Ī��Τ@�s�� (ex:����8��Ʀr)!");
				return true;
			} 
			if (isNaN(parseInt(pUserID.substring(0,8),10)) == true) //�T�w1-8�X�O�Ʀr
			{   alert("��J�L�Ī��Τ@�s�� (ex:����8��Ʀr)!");
				return true;
			} 
			for(ixI=0; ixI<=7; ixI++)//�M�����s�f
			{	tDiv=parseInt(parseInt(pUserID.substr(ixI,1),10)*parseInt(tStr.substr(ixI,1))/10);
				tMod=parseInt(parseInt(pUserID.substr(ixI,1),10)*parseInt(tStr.substr(ixI,1))%10);
				tSum=tSum+tDiv+tMod;
			}
			tSum=parseInt(tSum%10);
			         
			if ((tSum==0 || tSum==9) && pUserID.substr(6,1)=="7")
			{    return false;} //���T
			if (tSum==0)
			{   return false;} //���T
			else
			{	alert("��J�L�Ī��Τ@�s�� !"); //�����T
				return true;
			}
			break;   
	}
}
</script>
