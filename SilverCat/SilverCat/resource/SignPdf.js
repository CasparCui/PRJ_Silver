if (typeof SilverCat == "undefined") {
        var SilverCat = {};
}
SilverCat.sigField_cName = "SilverCatSignatureField";
SilverCat.sigField_cFieldType = "signature";

/**
* PDF�t�@�C���ɓd�q�������܂��B
* @param doc �d�q�����Ώۂ�PDF�h�L�������g�B
* @param jsonSignParam �d�q�����p�����[�^�B
*/
function SignToPdf(doc, jsonSignParam) {
  try {
    var field = AddSignatureField(doc, SilverCat.sigField_cName, SilverCat.sigField_cFieldType, jsonSignParam);
    if(field) {
      if(jsonSignParam.recipientPublicCerts == null) {
        EmbedSignToPdf(doc, field, jsonSignParam.password, jsonSignParam.digitalIdFilePath, jsonSignParam.securityPolicyName);
      } else {
        EmbedSignToPdf(doc, field, jsonSignParam.password, jsonSignParam.digitalIdFilePath, jsonSignParam.securityPolicyName, jsonSignParam.recipientPublicCerts);
      }
    }
    event.value = "0";
  } catch(e) {
    event.value = "Exception:" + e;
    console.println(e);
  }
}


/**
* PDF�h�L�������g�ɓd�q�����t�B�[���h��ǉ����܂��B
* @param doc �d�q�����Ώۂ�PDF�h�L�������g�B
* @param cName �d�q�����t�B�[���h���B
* @param cFieldType �t�B�[���h�^�C�v�B�d�q������ł̂ŁA"signature"�B
* @param jsonSignParam �d�q�����p�����[�^�B
* @return �d�q�����t�B�[���h�B
*/
function AddSignatureField(doc, cName, cFieldType, jsonSignParam) {
  // �t�B�[���h�̍��W��`
  // �����������_(x,y)��(0,0)
  // A4�c��PDF�h�L�������g��PageBox��get����ƁA
  // rect[0]=0,rect[1]=792,rect[2]=612,rect[3]=0
  // �ƂȂ�B
  var rectangle = doc.getPageBox( {nPage: 0} );
  
  rectangle[0] = jsonSignParam.x; // ����x���W�B�v���X�ŉE�ɍs���B
  rectangle[1] -= jsonSignParam.y; // ����y���W�B�}�C�i�X�ŉ��ɍs���B
  rectangle[2] = rectangle[0] + jsonSignParam.width; // ���B�v���X�ŉE�ɍs���B
  rectangle[3] = rectangle[1] - jsonSignParam.height; // ���B�}�C�i�X�ŉ��ɍs���B

  var field = null;
  try {
    // PDFDocument.addField()
    // cName:�t�B�[���h���BPDF�h�L�������g�̒��Ń_�u��Ȃ���΂悢�ł��傤�B
    // cFieldType:�t�B�[���h�^�C�v�B�d�q������ł̂ŁA"signature"�B
    // nPageNum:�y�[�W�ԍ��̓[������n�܂�B�[���Ȃ�΁A�擪�̃y�[�W�ɓd�q�����t�B�[���h��ǉ��B
    // aRect:�d�q�����t�B�[���h���W
    field = doc.addField(cName, cFieldType, 0, rectangle);
    // ���Ƃ���B
    field.hidden = false;
    // ����ΏۊO�Ƃ���B
    field.print = false;
  } catch (e) {
    throw e;
  }
  return field;
}

/**
* �d�q�����t�B�[���h�ɓd�q�����𖄂ߍ��݂܂��B
* @param sigField �d�q�����t�B�[���h�B
* @param pwd �d�q�ؖ����p�X���[�h�B
* @param did �d�q�ؖ����t�@�C���̃t���p�X�B
* @param policy �Z�L�����e�B�|���V�[���B
*/
EmbedSignToPdf = app.trustedFunction (
  function(doc, sigField, pwd, did, policy) {
    try {
      app.beginPriv();

      var sh = security.getHandler(security.PPKLiteHandler, false);
      sh.login(pwd, did);
      // ���̓d�q�ؖ����Ƀ��O�C������p�X���[�h�̃^�C���A�E�g�̓f�t�H���g�A�[���b�B
      // �f�t�H���g�l�̂܂܂��ƁA�d�q�����������̂悤�Ɍ����āA���͋�U�肷��悤�Ȃ̂ŁA
      // �悭�킩��Ȃ����[���b����Ȃ��l�A�Ⴆ��30�b���w��B
      sh.setPasswordTimeout(pwd, 30); 

      // �Z�L���e�B�|���V�[�ɂ��Z�L�����e�B�ݒ�F
      // Acrobat Pro XI�̏ꍇ�F
      // �u�\��(V)�v�u�c�[��(T)�v�Ńc�[���p�l����\���B
      // �u�ی�v�A�R�[�f�B�I���A�u�Í����v�A�R�[�f�B�I�����N���b�N���A
      // �u�Z�L�����e�B�|���V�[���Ǘ�(M)...�v���N���b�N�B
      // �u�Z�L�����e�B�|���V�[�̊Ǘ��v�_�C�A���O�œo�^�����Z�L�����e�B�[�|���V�[�̖��O��
      // �o�^���Ă������Z�L�����e�B�|���V�[�ɂ��APDF�t�@�C���ɃZ�L�����e�B��ݒ肷��B

      // �Z�L�����e�B�|���V�[�̌����B
      var options = { cHandler:security.StandardHandler };
      var policyArray = security.getSecurityPolicies( { oOptions: options } );
      var myPolicy = null;
      for( var i = 0; i < policyArray.length; i++) {
        if( policyArray[i].name == policy ) {
            myPolicy = policyArray[i];
            break;
        }
      }
      // �Z�L�����e�B�|���V�[�̐ݒ�B
      var res = doc.encryptUsingPolicy({ oPolicy: myPolicy , oHandler: sh, bUI: false } );
      if( res.errorCode != 0 ) {
        throw res.errorText;
      }

      // �d�q�������̐ݒ�B
      // mdp:"allowAll":���ʏ����B
      //     "allowNone","default","defaultAndComments":MDP�����B
      var sigInfo = {
        mdp: "allowNone"
      };
      // �d�q����������B
      sigField.signatureSign({oSig: sh, oInfo: sigInfo, bUI: false});

      app.endPriv();
    } catch (e) {
      throw e;
    }
  }
);


/**
* �d�q�����t�B�[���h�ɓd�q�����𖄂ߍ��݂܂��B
* @param sigField �d�q�����t�B�[���h�B
* @param pwd �d�q�ؖ����p�X���[�h�B
* @param did �d�q�ؖ����t�@�C���̃t���p�X�B
* @param policy �Z�L�����e�B�|���V�[���B
* @param recipientPublicCertFilePathArray ��M�҂̌��J���ؖ����t�@�C���t���p�X�̔z��B
*/
EmbedSignToPdf = app.trustedFunction (
  function(doc, sigField, pwd, did, policy, recipientPublicCertFilePathArray) {
    try {
      app.beginPriv();

      var sh = security.getHandler(security.PPKLiteHandler, false);
      sh.login(pwd, did);
      // ���̓d�q�ؖ����Ƀ��O�C������p�X���[�h�̃^�C���A�E�g�̓f�t�H���g�A�[���b�B
      // �f�t�H���g�l�̂܂܂��ƁA�d�q�����������̂悤�Ɍ����āA���͋�U�肷��悤�Ȃ̂ŁA
      // �悭�킩��Ȃ����[���b����Ȃ��l�A�Ⴆ��30�b���w��B
      sh.setPasswordTimeout(pwd, 30); 

      // ���J���ɂ��Z�L�����e�B�ݒ�:
      // PDF�t�@�C����M�҂̌��J���ňÍ�������B
      // PDF�t�@�C�����J���ɂ͏�L���J���̃y�A�ł���閧���łȂ��ƊJ���Ȃ��B
      // ��M�҂�Adobe Reader��Adobe Pro�Ɏ����̓d�q�ؖ�����o�^���āA�J�����ƂɂȂ�B
/*
      // ��M�҂̌��J���ؖ����̎�荞��(X509.3)
      var tarou = security.importFromFile("Certificate", "C:/temp/pub_tarou.cer" );
      var hanako = security.importFromFile("Certificate", "C:/temp/pub_hanako.cer" );
      // ��M�҃O���[�v���ɃA�N�Z�X������ݒ�
      var group1 = { userEntities: [ {certificates: [tarou] } ], permissions: {allowAll: true} };
      var group2 = { userEntities: [ {certificates: [hanako] } ], permissions: {allowChanges: "fillAndSign"} };
      // ��M�҂̌��J���ňÍ���
      doc.encryptForRecipients({ oGroups: [group1, group2] , bMetaData: true } );
*/
      var recipients = new Array();
      for each (var certPath in recipientPublicCertFilePathArray) {
        var cert = security.importFromFile("Certificate", certPath);
        recipients.push({certificates:[cert]});
      }

      var group = { userEntities: recipients, permissions: {allowAll: true} };
      doc.encryptForRecipients({ oGroups: group, bMetaData: true, bUI: false } );

      // �d�q�������̐ݒ�B
      // mdp:"allowAll":���ʏ����B
      //     "allowNone","default","defaultAndComments":MDP�����B
      var sigInfo = {
        mdp: "allowNone"
      };
      // �d�q����������B
      sigField.signatureSign({oSig: sh, oInfo: sigInfo, bUI: false});

      app.endPriv();
    } catch (e) {
      throw e;
    }
  }
);

