# example
This repository is for practising the GitHub Flow


```javascript
function getUsers_TRIGGER() {

  var start = Date.now();

  var sheet_users = SpreadsheetApp.getActive().getSheetByName( STR.sheet_name_users );

  var user_properties    = PropertiesService.getUserProperties();
  var continuation_token =  user_properties.getProperty( 'FILES_CONTINUATION_TOKEN' );
  var uniques_save       =  user_properties.getProperty( 'UNIQUE_ACCOUNTS'          );
  var processed_count    = +user_properties.getProperty( 'PROCESSED_COUNT'          );

  var uniques = uniques_save ? JSON.parse( uniques_save ) : {};

  var files   = continuation_token == null ?
                DriveApp.searchFiles( `'me' in owners and trashed=false and visibility='limited'` ) :
                DriveApp.continueFileIterator( continuation_token );
  
  
  // ОБРАБОТКА
  var obj = processFiles( files, uniques, processed_count, start );

  SpreadsheetApp.getActive().getSheetByName( STR.sheet_name_log )
    .appendRow( [ null, null, 'INFO', `Обработано: ${obj.processed_count}`, obj.files.hasNext() ? 'Продолжаем...' : 'КОНЕЦ!' ] );


  // ЗАВЕРШЕНИЕ
  if ( obj.files.hasNext() ) { // Не успели обработать все файлы

    user_properties.setProperties({
      'FILES_CONTINUATION_TOKEN' : obj.files.getContinuationToken(),
               'UNIQUE_ACCOUNTS' : JSON.stringify( obj.uniques ),
               'PROCESSED_COUNT' : obj.processed_count
    });

    deleteTriggers();
    ScriptApp.newTrigger( LIB + 'getUsers_TRIGGER' ).timeBased().after(1).create(); /*TRIG*/
  
  } else {

    deleteSettings(); // Удаляем триггеры и свойства

    // ВЫВОД
    
    // Ни одного файла с личным доступом???
    if ( !Object.keys( obj.uniques ).length ) { status( STR.msg_users_done_wtf ); return; } /*MSG*/ //RETURRRRRRNNNNNNNNNN
 

    status( STR.msg_users_done ); /*MSG*/

    var values = Object.entries( obj.uniques )
      .sort( ( a, b ) => a[1].localeCompare( b[1], undefined, { sensitivity : 'base' } ) ) // сортим а-я кейс инсенситив по емейлу
      .filter( e => e[0] ) // фильтруем от пустых Account deleted (удалённые гуглом аккаунты, у которых нет ни имени, ни почты)
      .map( entry => [ entry[1], entry[0], false ] ); // имя, ящик, пустой чекбокс = false

    var index = 3 + values.length - 1;

    sheet_users.getRange( 'A3:C' + index ).setValues( values );
    sheet_users.getRange( 'C3:C' + index ).insertCheckboxes();

  }
}
```
