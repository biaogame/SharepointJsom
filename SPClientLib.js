import $ from 'jquery';

let SPClientLib = {

}

/**
 * @description 获取item的授权的组或用户
 */
SPClientLib.getRoleAssignments = {
  url: "/_api/web/lists/getById('{0}')/items/getById({1})/roleassignments",
  type: 'GET',
  headers: {
    'accept': 'application/json;odata=verbose',
    "X-RequestDigest":document.getElementById('__REQUESTDIGEST')==null?'': document.getElementById('__REQUESTDIGEST').value
  }
}
/**
 * @description 获取站点的所有用户
 */
SPClientLib.getSiteUsers = {
  url: '/_api/web/siteusers',
  type: 'GET',
  headers: {
    'accept': 'application/json;odata=verbose',
    "X-RequestDigest": document.getElementById('__REQUESTDIGEST')==null?'':document.getElementById('__REQUESTDIGEST').value
  }
}

/**
 * @description 获取站点的所有用户组
 */
SPClientLib.getSiteGroups = {
  url: '/_api/web/sitegroups',
  type: 'GET',
  headers: {
    'accept': 'application/json;odata=verbose',
    "X-RequestDigest":document.getElementById('__REQUESTDIGEST')==null?'': document.getElementById('__REQUESTDIGEST').value
  }
}
/**
 * 
 * @param {*} listId 
 * @param {*} itemid 
 */
SPClientLib.getRoleAssignmentsFun=function(listId,itemid,_callback){
  var url = "/_api/web/lists/getById('{0}')/items/getById({1})/roleassignments";
  $.ajax({
    url: url.format(listId, itemid),
    method: "GET",
    headers: {
      "ACCEPT": "application/json;odata=verbose",
      'X-RequestDigest': document.getElementById('__REQUESTDIGEST')==null?'':document.getElementById('__REQUESTDIGEST').value,
    },
    success: function (a, b) {
      _callback(a,b);
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value,'fail');
    }
  });
}



/**
 * @description 从item中移除用户或组
 */
SPClientLib.removeRoleassign = function (listId, itemid, principalid, _callback) {
  var url = "/_api/web/lists/getById('{0}')/items/getById('{1}')/roleassignments/getbyprincipalid('{2}')";
  $.ajax({
    url: url.format(listId, itemid, principalid),
    method: "POST",
    headers: {
      "ACCEPT": "application/json;odata=verbose",
      'X-RequestDigest': document.getElementById('__REQUESTDIGEST')==null?'':document.getElementById('__REQUESTDIGEST').value,
      'X-HTTP-Method': 'DELETE'
    },
    success: function (a, b) {
      _callback(b);
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value);
    }
  });
}
/**
 * @description 添加权限 1073741826读取 1073741927仅查看 Full Control: 1073741829
 */
SPClientLib.addRoleAssignment = function (listId, itemid, principalid, _callback) {
  var url = "/_api/web/lists/getById('{0}')/items/getById({1})/roleassignments/addroleassignment(principalid={2},roledefid=1073741829)";
  $.ajax({
    url: url.format(listId, itemid, principalid),
    method: "POST",
    headers: {
      "ACCEPT": "application/json;odata=verbose",
      'X-RequestDigest': document.getElementById('__REQUESTDIGEST')==null?'':document.getElementById('__REQUESTDIGEST').value
    },
    success: function (a, b) {
      _callback(b);
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value);
    }
  });
}


/**
 * @description 删除继承权限，使用独立权限
 * copyroleassignments=true
 */
SPClientLib.breakRoleInheritance = function (listId, itemid, _callback) {
  var url = "/_api/web/lists/getById('{0}')/items/getById({1})/breakroleinheritance(copyroleassignments=true,clearsubscopes=true)";
  $.ajax({
    url: url.format(listId, itemid),
    method: "POST",
    headers: {
      "ACCEPT": "application/json;odata=verbose",
      'X-RequestDigest': document.getElementById('__REQUESTDIGEST')==null?'':document.getElementById('__REQUESTDIGEST').value
    },
    success: function (a, b) {
      _callback(b);
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value);
    }
  });

}

/**
 * 授权，并删除列表的继承权限
 * @param {*} listId 
 * @param {*} itemid 
 * @param {*} principalid 
 * @param {*} _callback 
 */
SPClientLib.addRoleAssignmentAndBreakRoleInheritance = function (listId, itemid, principalid, _callback) {
  SPClientLib.breakRoleInheritance(listId, itemid, function (d) {
    if (d == 'success') {
      SPClientLib.addRoleAssignment(listId, itemid, principalid, function (d2) {
        _callback(d2);
      })
    }
  })
}
/**
 * @description 添加item
 * @param  listtitle 
 * @param {*} newItem 
 * @param {*} _callback 
 */
SPClientLib.addItem = function (listtitle, newItem, _callback) {
  /*
  newItem['__metadata'] = {
    "type": "SP.Data.List1ListItem"
  }*/
  var requestBody = JSON.stringify(newItem);
  var requestHeaders = {
    "ACCEPT": "application/json;odata=verbose",
    "X-RequestDigest": document.getElementById('__REQUESTDIGEST').value,
  }
  $.ajax({
    url: "/_api/web/lists/getbytitle('" + listtitle + "')/items",
    method: "POST",
    contentType: "application/json;odata=verbose",
    headers: requestHeaders,
    data: requestBody,
    success: function (data, b) {
      _callback(data, b);
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value,'fail');
    }
  });
}

/**
 * @description 编辑item
 * @param  listtitle 
 * @param {*} newItem 
 * @param {*} _callback 
 */
SPClientLib.editItem = function (listtitle, newItem, ID, _callback) {
  var requestBody = JSON.stringify(newItem);
  var requestHeaders = {
    "ACCEPT": "application/json;odata=verbose",
    "X-RequestDigest": document.getElementById('__REQUESTDIGEST').value,
    "X-HTTP-Method": "MERGE",
    "If-Match": "*"
  }

  $.ajax({
    url: "/_api/web/lists/getbytitle('" + listtitle + "')/items('" + ID + "')",
    method: "POST",
    contentType: "application/json;odata=verbose",
    headers: requestHeaders,
    data: requestBody,
    success: function () {
      newItem['Id']=ID;
      _callback({d:newItem}, 'success');
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value,'fail');
    }

  });
}

/**
 * @description 删除item
 * @param {*} listtitle 
 * @param {*} itemid 
 * @param {*} _callback 
 */
SPClientLib.delItem = function (listtitle, itemid, _callback) {
  var requestHeaders = {
    "ACCEPT": "application/json;odata=verbose",
    "X-RequestDigest": document.getElementById('__REQUESTDIGEST').value,
    "X-Http-Method": "DELETE",
    "IF-MATCH": "*"
  }
  $.ajax({
    url: "/_api/web/lists/getbytitle('" + listtitle + "')/items(" + itemid + ")",
    method: "POST",
    contentType: "application/json;odata=verbose",
    headers: requestHeaders,
    success: function () {
      _callback('success');
    },
    error: function (jqXHR, textStatus, errorThrown) {
      _callback(jqXHR.responseJSON.error.message.value,'fail');
    }

  });
}

/**
 * @description 获取当前站点id
 */
SPClientLib.GetSiteId=function(_callback){
  var curContext = new SP.ClientContext();
  var curSite = curContext.get_site();
  curContext.load(curSite);
  curContext.executeQueryAsync(Function.createDelegate(this, onCreateSucceeded), Function.createDelegate(this, onCreateFailed));
  function onCreateSucceeded(sender, args) {
      //var siteId = curSite.get_id().toString();
      _callback(curSite.get_id().toString());
  }
  function onCreateFailed(sender, args) {
    _callback('-1');
  }
}


/**
 * @description 判断用户是否web的管理员提供方法 get_loginName() get_id() get_title() get_email()
 * @param {function} _callback 回调函数 返回真假
 */
SPClientLib.isUserHostWebAdmin = function (_callback, isInfo = false) {
  ExecuteOrDelayUntilScriptLoaded(function () { //先要判断sp.js是否已经加载完成
    var clientContext = new SP.ClientContext(_spPageContextInfo.siteAbsoluteUrl);
    var web = clientContext.get_web();
    var currentUser = web.get_currentUser();
    clientContext.load(web);
    clientContext.load(currentUser);

    clientContext.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
        if (isInfo) {
            _callback(currentUser);
        } else {
          var isSiteAdmin = currentUser.get_isSiteAdmin();
          _callback(isSiteAdmin);
        }

      }),
      Function.createDelegate(this, function (sender, args) {}));
  }, 'sp.js')
}

/**
 * @description 获取下拉选项的字段值
 * @param {function} _callback 回调函数 返回真假
 */
SPClientLib.GetColumnChoice = function (listname, fieldname, _callback) {
  ExecuteOrDelayUntilScriptLoaded(function () { //先要判断sp.js是否已经加载完成
      var context = new SP.ClientContext.get_current();
      var web = context.get_web();
      var testList = web.get_lists().getByTitle(listname);
      var columnchoice = context.castTo(testList.get_fields().getByInternalNameOrTitle(fieldname), SP.FieldChoice);
      context.load(columnchoice);
      context.executeQueryAsync(onSuccessMethod, onFailureMethod);
      function onSuccessMethod(sender, args) {
          var choices = columnchoice.get_choices();
          _callback(choices);
      }

      function onFailureMethod(sender, args) {
          _callback('-1');
      }
  }, 'sp.js')
}



/**
 * @description 使用JSOM对象模型获取列表分页数据
 * @param {string} q_title 列表的标题
 * @param {string} q_caml caml查询语句
 * @param {string} q_pagingInfo 分页信息 
 * @param {function} _callback 回调函数
 */
SPClientLib.GetListByPage = function (q_title, q_caml, q_pagingInfo, _callback) {
  q_pagingInfo = q_pagingInfo || 'PagePrev=True&Paged=TRUE';
  ExecuteOrDelayUntilScriptLoaded(function () { //先要判断sp.js是否已经加载完成
    var clientContext = new SP.ClientContext(_spPageContextInfo.siteAbsoluteUrl);
    var oList = clientContext.get_web().get_lists().getByTitle(q_title);
    var camlQuery = new SP.CamlQuery();
    camlQuery.set_viewXml(q_caml);
    var position = new SP.ListItemCollectionPosition();
    position.set_pagingInfo(q_pagingInfo);
    camlQuery.set_listItemCollectionPosition(position);
    var collListItem = oList.getItems(camlQuery);

    clientContext.load(collListItem);


    clientContext.executeQueryAsync(
      Function.createDelegate(this, function (sender, args) {
        var result = {};
        result.ListItemCollection = collListItem.getEnumerator(); //返回列表
        result.pageCount = collListItem.get_count();
        result.PagingInfo = collListItem.get_listItemCollectionPosition() != null ? collListItem.get_listItemCollectionPosition().get_pagingInfo() : null; //返回下一页分页信息
        _callback(result);
      }),
      Function.createDelegate(this, function (sender, args) {
        console.log(sender);
        console.log(args);
        _callback(-1);

      })
    );
  }, "sp.js");
}

/**
 * @description 根据数据生成camlxml数组形式[{'Type':'Text','Name':'Title','Val':'测试','Opera':'Contains'}]
 *Type字段类型(Text Number Integer DateTime(如果不填IncludeTimeValue 默认True))
 *现在只有AND链接语句 若有OR语句请不要使用该方法
 */
SPClientLib.BuildeCaml = function (itemArr) {
  var result = '';
  if (itemArr.length <= 3) {
    if (itemArr.length == 3) {
      result += '<And><And>';
      result += SPClientLib.BuildeQuery(itemArr[0]);
      result += SPClientLib.BuildeQuery(itemArr[1]);
      result += '</And>';
      result += SPClientLib.BuildeQuery(itemArr[2]);
      result += '</And>';
    } else if (itemArr.length == 2) {
      result += '<And>';
      result += SPClientLib.BuildeQuery(itemArr[0]);
      result += SPClientLib.BuildeQuery(itemArr[1]);
      result += '</And>';
    } else if (itemArr.length == 1) {
      result += SPClientLib.BuildeQuery(itemArr[0]);
    }
  } else {
    var str = '';
    str += '<And>';
    str += SPClientLib.BuildeQuery(itemArr[0]);
    str += SPClientLib.BuildeQuery(itemArr[1]);
    str += '</And>';
    for (var i = 2; i < itemArr.length; i++) {
      str = SPClientLib.CombineQuery(str, itemArr[i], i);
    }
    result = str;
  }
  if (result != '') {
    result = '<Where>' + result + '</Where>';
  }
  return result;
}
SPClientLib.BuildeQuery = function (item) {

  var str = '';
  var looupvalue = '';
  if (item.hasOwnProperty('LookupId') && item.LookupId == true) { //判断是否查阅项
    looupvalue = ' LookupId="TRUE" ';
  }
  if (item.Type == 'DateTime') {
    var IncludeTimeValue = 'True';
    if (item.hasOwnProperty('IncludeTimeValue')) {
      IncludeTimeValue = item.IncludeTimeValue;
    }
    str = '<' + item.Opera + '><FieldRef Name="' + item.Name + '" ' + looupvalue + ' /><Value IncludeTimeValue="' + IncludeTimeValue + '" Type="' + item.Type + '">' + item.Val + '</Value></' + item.Opera + '>';
  } else {
    if (item.Opera == 'In') { //如果条件是In
      str += '<In>';
      str += '<FieldRef Name="' + item.Name + '" />';
      str += '<Values>';
      var valsArr = item.Val.split(',');
      for (var i = 0; i < valsArr.length; i++) { //效果如下 <Values><Value Type="Integer">1</Value><Value Type="Integer">4</Value></Values>
        if (valsArr[i] != '') {
          str += '<Value Type="' + item.Type + '">' + valsArr[i] + '</Value>';
        }
      }
      str += '</Values>';
      str += '</In>';
    } else if (item.Opera == 'IsNotNull' || item.Opera == 'NotNull') {
      str = '<' + item.Opera + '><FieldRef Name="' + item.Name + '" ' + looupvalue + ' /></' + item.Opera + '>';
    } else {
      str = '<' + item.Opera + '><FieldRef Name="' + item.Name + '" ' + looupvalue + ' /><Value Type="' + item.Type + '">' + item.Val + '</Value></' + item.Opera + '>';
    }

  }
  return str;
}

SPClientLib.CombineQuery = function (str, item, i) {
  str = '<And>' + str;
  str += SPClientLib.BuildeQuery(item);
  str += '</And>';
  return str;
}



SPClientLib.Alert = function () {
  alert('');
}
export default SPClientLib