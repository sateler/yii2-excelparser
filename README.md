# yii2-excelparser
Parse excel files using PHP_Excel into yii2 models
# Examples:
Model
```php
class ImportForm extends \yii\base\Model
{
    private static $fields = [
        'Name' => 'name',
        'Code' => 'code',
        'Subject' => 'subject',
    ];
    
    private static $required = [
        'Name',
        'Code',
    ];
    
    public $file;

    /**
     * @inheritdoc
     */
    public function rules()
    {
        return [
            ['file', 'setFileProp', 'skipOnEmpty' => false],
            [['file'], 'file', 'skipOnEmpty' => false],
        ];
    }

    /**
     * @inheritdoc
     */
    public function attributeLabels()
    {
        return [
            'file' => 'Archivo',
        ];
    }
    
    public function setFileProp() {
        $this->file = UploadedFile::getInstance($this, 'file');
    }
    
    public function import() {
        try {
            $parser = new ExcelParser([
                'fileName' => $this->file->tempName,
                //'worksheetName' => 'Sheet 1',
                'fields' => self::$fields,
                'requiredFields' => self::$required,
                'setNullValues' => false,
                'modelClass' => SomeModel::className(),
                // Use either modelClass or createObject...
                // createObject is called before parsing the row.
                'createObject' => function ($prevRow) use ($someData) {
                    $newObj = new SomeModel();
                    $newObj->parent_id = $someData->id;
                    if (!$prevRow) {
                        return $newObj;
                    }
                    // set some defualt data from previous created object
                    foreach(self::$copyFields as $field) {
                        $newObj->$field = $prevRow->$field;
                    }
                    return $newObj;
                },
                // Will save the data in an internal array, set to false for large datasets to save memory
                'saveData' => false,
                // Callback after object has been created and parsed
                'onObjectParsed' => function(SomeModel $data, $rowIndex) {
                    return $data->save();
                },
            ]);
        }
        if ($parser->getError()) {
            Yii::error("ExcelParser Error: " . $parser->getError());
            $this->addError('file', 'Archivo con formato inválido');
            return false;
        }
        
        // If 'savedata' is set to true, then get the data:
        $allData = $parser->getData();
        foreach($allData as $i => $data) {
            if (!$data->save()) {
                $this->addError('file', "Error en fila $i: " . implode("\n", $data->getFirstErrors()));
                return false;
            }
            unset($allData[$i]);
        }
        return true;
    }
}
```
Controller
```php

    public function actionImport()
    {
        $form = new ImportForm;
        
        if ($form->load(Yii::$app->request->post()) && $form->validate()) {
            if($form->import()) {
                Yii::$app->session->setFlash('success', 'Success');
                return $this->redirect(['import']);
            }
            else {
                Yii::$app->session->setFlash('error', 'Error');
            }
        }
        
        return $this->render('import', [
            'model' => $form,
        ]);
    }
````
