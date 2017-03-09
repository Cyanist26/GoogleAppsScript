/**
 * 在給定範圍中尋找指定內容，支持兩種搜索模式，返回搜索結果在輸入範圍中的行、列偏移量。
 *
 * @param {FALSE} mode 搜索模式，FALSE使用正則表達式匹配，TRUE為完全匹配。
 * @param {abc} content 需要搜索的內容。
 * @param {data} data 需要搜索的範圍。
 * @return 返回搜索得到的地址。
 */
function getIndexByContent(mode, content, data){
  try
  {
    if( arguments.length != 3 ) throw "ERR_ARGS";
    if( arguments[0] != true && arguments[0] != false ) throw "#ERR_MODE";
    
    var result = [];
    
    if( mode )
      for( var i = 0; i < data.length; i++ )
      { 
        for( var j = 0; j < data[i].length; j++ )
        {
          if( data[i][j] == content )
            result.push( [i + 1,j + 1] );
        }
      }
    else
    {
      var reg = new RegExp(content,i);
      for( var i = 0; i < data.length; i++ )
      { 
        for( var j = 0; j < data[i].length; j++ )
        {
          if( reg.test(data[i][j].toString()) )
            result.push( [i + 1,j + 1] );
        }
      }
    }
    
    if( result.length == 0 )
      throw "#ERR_NOTFOUND";
    else
      return result;
  }
  catch(e)
  {
    SpreadsheetApp.getUi().alert(e);
    return null;
  }
}


/**
 *一個簡易字典對象。
 *包含containsKey,push,remove,getValue,setValue,getLength,toString共7個方法。
 */
function dictionary() {
  try
  {
    var _dict = new Object();
    var _length = 0;
    
     /**
     * 檢查key是否已存在。
     *
     * @param {key} key key。
     * @return true or false.
     */
    this.containsKey = function(key) {
      var isContained = false;
      for( var e in _dict ){
        if( e == key ){
          isContained =true;
          break;
        }
      }
      return isContained;
    }
    
    /**
     * 新增key-value對。
     *
     * @param {key} key key。
     * @param {value} value value。
     * @return 當前字典鍵值對數。
     */
    this.push = function(key, value) {
      if( !this.containsKey(key) ){
        _dict[key] = value;
        _length ++;
        return _length;
      }
      else
        return false;
    }
    
    /**
     * 刪除key-value對。
     *
     * @param {key} key key。
     * @return true or false.
     */
    this.remove = function(key) {
      if(this.containsKey(key)){ 	 	
        delete _dict[key];
        _length --;
        return true;
      }
      else
        return false;
    }
    
    /**
     * 根據key返回對應value。
     *
     * @param {key} key key。
     * @return value or false.
     */
    this.getValue = function(key) {
      if(this.containsKey(key)){
        return _dict[key];
      }
      else
        return null;
    }
    
    /**
     * 根據key更改對應value。
     *
     * @param {key} key key。
     * @param {value} value value。
     * @return true or false.
     */
    this.setValue = function(key, value) {
      if( this.containsKey(key) )
      {
        _dict[key] = value;
        return true;
      }
      else
        return false;
    }
    
    /**
     * 返回鍵值對數。
     *
     * @return 鍵值對數。
     */
    this.getLength = function() {
      return _length;
    }
    
    /**
     * 返回JSON格式字符串。
     *
     * @return JSON格式字符串。
     */
    this.toString = function(){ 	 	
      var result = JSON.stringify(_dict); 	 	
      return result; 	 	
    }
  }
  catch(e)
  {
    throw e;
  }
}

/**
 * 計算輸入值的MD5摘要。
 *
 * @param {key} key 輸入的字符串值。
 * @return 返回字符串形式的MD5摘要。
 */
function computeMD5(key){
  try
  {
    var contentDigestByte = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key, Utilities.Charset.US_ASCII);
    var contentDigest = '';
    for( var i = 0; i < contentDigestByte.length; i++ )
    {
      var byte = contentDigestByte[i];
      if (byte < 0)  byte += 256;
      var byteStr = byte.toString(16);
      if (byteStr.length == 1) byteStr = '0'+byteStr;
      contentDigest += byteStr;
    }
    
    return contentDigest;
  }
  catch(e)
  {
    SpreadsheetApp.getUi().alert(e);
  }
}