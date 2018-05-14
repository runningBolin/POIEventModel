# POIEventModel
使用POI的EventModel解析07版及以上的excel文件。poi的event model对内存和cpu资源的占用都极小，解析效率也比user model强很多，5w+数据解析时间少于一分半。

提供两种读取sheet页内数据的方式：1，一次性解析出指定sheet页内数据行，提供数据出口；2，逐行获取sheet页内数据行，基于BlockingQueue阻塞队列实现，提供逐行获取数据的出口；

#### 方式1示例:
    SheetHandler sheetHandler = new SheetHandler();
    WorkbookHandler handler = new WorkbookHandler(sheetHandler);
    handler.parseSheet("F:\\CBT_Eligible_Items_20170725.xlsx", "Pivot_Super Department");
    List<Map<Integer, String>> sheetDataList = sheetHandler.getDataList();
    if(sheetDataList != null && !sheetDataList.isEmpty()){
      for ( Map<Integer, String> map : sheetDataList ) {
        for ( int i = 0; i < map.size(); i++ ) {
          String cellValue = map.get(i+1);
          System.out.println("cellValue:"+ cellValue);
        }
      }
    }
    

#### 方式2示例

    QueueSheetHandler sheetHandler = new QueueSheetHandler();
    WorkbookHandler handler = new WorkbookHandler(sheetHandler);
    handler.parseSheet("F:\\CBT_Eligible_Items_20170725.xlsx", "CBT_Eligible_Items_20170725");
    //取数据需要判断当前解析是否已结束
    while(!sheetHandler.finished()){
      HashMap<Integer, String> rowData = sheetHandler.readRowData();
      //使用poi event model解析excel文档时，空行也会被解析到，需要判断当前获取的行数据是否是空集合
      if(rowData != null && rowData.isEmpty()){
          System.out.println("rowData:"+ rowData.toString());
      }
    }
    
