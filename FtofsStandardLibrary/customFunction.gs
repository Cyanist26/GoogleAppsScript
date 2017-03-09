 /**
 * 合併顯示兩個或多個輸入範圍，支持單列或多列輸入。不限制參數個數。
 *
 * @param {FALSE} MODE 合併模式，FALSE為保留空行，TRUE為剔除空行。
 * @param {A1:B10} RANGE1 需要合併的第一個範圍。
 * @param {array} RANGE2 需要合併的第二個範圍。
 * @return 返回合併後的範圍。
 * @customfunction
 */
//function MERGEMULTIRANGES(MODE,RANGE1,RANGE2){
//  var args = [];
//  for(var i = 0; i < arguments.length; i++)
//    args.push(arguments[i]);
//  return FtofsStandardLibrary.MERGEMULTIRANGES(args);
//}
//function MERGEMULTIRANGES(MODE,RANGE1,RANGE2){
//  return FtofsStandardLibrary.MERGEMULTIRANGES(arguments);
//}
function MERGEMULTIRANGES()
{
  try
  {
    if( arguments[0].length <= 2 ) throw "ERR_ARGS";    //參數個數不足
    if( arguments[0][0] != true && arguments[0][0] != false ) throw "#ERR_MODE";    //模式參數錯誤
    
    var result = [];
    
    /** 合拼輸入範圍 */
    for (var i = 1; i < arguments[0].length; i++) 
    {
      result = result.concat(arguments[0][i]);
    }
    
    /** 刪除模式下刪除空行 */
    if (arguments[0][0])
    {
      for (var row = 0; row < result.length; row++)
      {
        for (var col = 0; col < result[row].length; col++)
        {
          if (result[row][col] != "") break;
          if (col == (result[row].length - 1) && result[row][col] == "")
          {
            result.splice(row,1);
            row--;
            break;
          }
        }
      }
    }
    return result;
  }
  catch(e)
  {
    return e;
  }
}

 /**
 * 檢查輸入的兩組數據是否一致，僅支持單列或單行數據。
 * 支持排序輸出和亂序輸出兩種模式。
 *
 * @param {A1:A10} RANGE1 包含全部數據的範圍。
 * @param {B1:B10} RANGE2 缺少部份數據的範圍。
 * @param {B1:B10} MODE 輸出模式，TRUE為排序，FALSE為亂序，可省略，默認為TRUE。
 * @return 返回添加空行后的範圍。
 * @customfunction
 */
//function FINDLOSTDATA(RANGE1, RANGE2, MODE){
//  return FtofsStandardLibrary.FINDLOSTDATA(RANGE1, RANGE2, MODE);
//}

function FINDLOSTDATA(RANGE1, RANGE2, MODE) {
  var index1;
  var index2;
  var result = [];
  if( arguments.length == 2 || MODE == undefined) MODE = true;
  
  try
  {
    if( arguments.length != 2 && arguments.length != 3 ) throw "ERR_ARGS";    //參數個數錯誤
    if( RANGE1 == undefined || RANGE2 == undefined || MODE == undefined) throw "ERR_ARG";    //參數輸入錯誤
    if( arguments.length == 3 && MODE != true && MODE != false ) throw "#ERR_MODE";    //模式參數錯誤
    if( RANGE1[0].length > 1 || RANGE2[0].length > 1) throw "#ERR_AREA_COL";    //輸入範圍錯誤
    
    if( MODE )
    /** sort mode */
    {   
      RANGE1 = RANGE1.filter(function(input){return input != ""}).sort(function(a,b){return a-b;})
      RANGE2 = RANGE2.filter(function(input){return input != ""}).sort(function(a,b){return a-b;})
      if( RANGE1.length <= RANGE2.length ) throw "#ERR_AREA_ROW";    //輸入範圍數據錯誤
      
      for( index1 = 0, index2 = 0; index1 < RANGE1.length && index2 < RANGE2.length; index1++ )
      {
        if( RANGE1[index1].toString() == RANGE2[index2].toString() )
        {
          result.push(RANGE1[index1]);
          index2++;
        }
        else
        {
          if( ( index1 - index2) > ( RANGE1.length - RANGE2.length ) ) throw "#ERR_DATA_AT"+( index2 + 1 );    //小範圍中出現大範圍中未包含的數據
          result.push("Lost_data_"+RANGE1[index1]);
        }
      }   
      /** 大範圍尾部數據丟失 */
      if (index1 < RANGE1.length && index2 >= RANGE2.length)
        for( ; index1 < RANGE1.length; index1++)
          result.push("Lost_data_"+RANGE1[index1]);
      
      return result;
    }
    else
    /** unsort mode*/
    {
      RANGE1 = RANGE1.filter(function(input){return input != ""});
      RANGE2 = RANGE2.filter(function(input){return input != ""});
      if( RANGE1.length <= RANGE2.length ) throw "#ERR_AREA_ROW";    //輸入範圍數據錯誤
      
      index2 = -1;
      
      /** 谷歌後台indexOf()函數出錯，覆蓋該函數 */
      Array.prototype.indexOf = function(a){
        for (var i = 0; i < this.length; i++ )
          if( this[i].toString() == a.toString() ) return i;
        return -1;}
      
      for( index1 = 0; index1 < RANGE1.length; index1++ )
      {
        index2 = RANGE2.indexOf(RANGE1[index1]);
        if( index2 > -1 )
        {
          RANGE2.splice(index2,1);
          result.push(RANGE1[index1]);
        }
        else
        {
          result.push("Lost_data_"+RANGE1[index1]);
        }
      }
      return result;
    }
  }
  catch(e)
  {
    result.push(e);
    return result;
  }
}

/**
 * 計算輸入值的MD5摘要，不限制參數個數。
 *
 * @param {Content} KEY 輸入的值（一個或多個）。
 * @return 返回字符串形式的MD5摘要。
 * @customfunction
 */
//function toMD5(KEY){
//  try{
//  var args = new Array(arguments.length);
//  for(var i = 0; i < args.length; i++)
//    args[i] = arguments[i];
//    
//  var result = FtofsStandardLibrary.computeMD5(args.join(""));
//    
//  return result;
//    
//  }
//  catch(e)
//  {
//    return e.toString();
//  }
//}