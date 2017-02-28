# yii2-excelparser
Parse excel files using PHP_Excel into yii2 models
# Excamples:
Model
```
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
                'fields' => self::$fields,
                'requiredFields' => self::$required,
                'setNullValues' => false,
                'modelClass' => 'common\models\SomeModel',
                // Use either modelClass or createObject...
                'createObject' => function ($prevRow) use ($import) {
                    $newObj = new SomeModel();
                    $newObj->import_id = $import->id;
                    if (!$prevRow) {
                        return $newObj;
                    }
                    foreach(self::$copyFields as $field) {
                        $newObj->$field = $prevRow->$field;
                    }
                    return $newObj;
                }
            ]);
        }
        catch (\Exception $e) {
            Yii::error("Error: " . $e->getMessage() . "\n" . $e->getTraceAsString());
            $this->addError('file', 'Archivo con formato inválido');
            return false;
        }
        if ($parser->getError()) {
            Yii::error("ExcelParser Error: " . $parser->getError());
            $this->addError('file', 'Archivo con formato inválido');
            return false;
        }
        
        $allData = $parser->getData();
        unset($parser);
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
```

    public function actionImport()
    {
        $form = new ImportForm;
        
        if ($form->load(Yii::$app->request->post()) && $form->validate()) {
            if($form->import()) {
                Yii::$app->session->setFlash('success', 'Success');
            }
            else {
                Yii::$app->session->setFlash('error', 'Error');
            }
            return $this->redirect(['import']);
        }
        
        return $this->render('import', [
            'model' => $form,
        ]);
    }
````
