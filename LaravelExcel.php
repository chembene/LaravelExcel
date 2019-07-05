 
use Excel;
use PHPExcel_Worksheet_Drawing;
use PHPExcel_Style_Fill;
use Doctrine\DBAL\Schema\Column;
use Intervention\Image\Gd\Font;
use Maatwebsite\Excel\Classes\PHPExcel;
use Maatwebsite\Excel\Writers\CellWriter;
 
 public functio export()
 {
    $empresa =DB::table('cust_empresa')->where('id', '=',  $this->empresa)
           ->select('id','nome', 'email', 'nuit', 'caixa_postal', 'contacto_1', 'contacto_2', 'morada','logo')->first(); 

       $data['periodo']= DB::table('periodos_pagamento')->where('id','=',$id)->first();   
       $data1 = DB::table('recibos')
                        ->leftjoin('funcionarios','funcionarios.id','=','recibos.funcionario_id')
                        ->leftjoin('categorias','categorias.id','=','funcionarios.categoria_id')
                  ->where('recibos.periodo_id',$id)
               ->select('recibos.*','funcionarios.nome','funcionarios.id as codFuncionario','funcionarios.salario_base','funcionarios.codigo_interno as codigo','categorias.nome as categoria')   
                 ->get();

        $logo= $empresa->logo;          
        $data = Funcionario::get(['nome','codigo_interno','genero', 'tipo_documento','nr_documento','nacionalidade','naturalidade', 'contacto', 'contacto_alternativo', 'categoria_id','empresa_id'])->toArray();


        $array = json_decode(json_encode($data1), true);


        $dadosFolha = array();
       
       

          foreach ($data1 as $key => $value) {
         
              $dadosFolha[$key]['Codigo Interno'] = $value->codigo;
              $dadosFolha[$key]['Nome'] = $value->nome;
              $dadosFolha[$key]['Categoria'] = $value->categoria;
              $dadosFolha[$key]['Salario'] = $value->salario_base;
              $dadosFolha[$key]['Bonus Fixos'] = $value->sub_fixos;
              $dadosFolha[$key]['Bonus variaveis'] = $value->sub_variaveis;
              $dadosFolha[$key]['Horas'] = $value->horas_extras;
              $dadosFolha[$key]['Faltas'] = $value->horas_falta;
              $dadosFolha[$key]['Total remuneracoes'] = ($value->horas_extras+$value->sub_fixos+$value->sub_variaveis+$value->salario_base+$value->horas_falta);
              $dadosFolha[$key]['IRPS'] = $value->irps;
              $dadosFolha[$key]['INSS'] = $value->inss;
              $dadosFolha[$key]['Outros'] =$value->irps+$value->inss;
              $dadosFolha[$key]['Total'] =$value->irps+$value->inss;
              $dadosFolha[$key]['Salario salario_liquido'] = $value->salario_liquido;


          }

    return Excel::create('Folha de Salarios', function($excel) use ($dadosFolha,$empresa) {
      
        $excel->sheet('primeira', function($sheet) use ( $dadosFolha, $empresa)
        {

           

          
            $objDrawing = new PHPExcel_Worksheet_Drawing;
            $objDrawing->setPath(public_path('Anexos/logo_empresas\\'.$empresa->logo)); //your image path
            $objDrawing->setCoordinates('A2'); //posision da image
            $objDrawing->setHeight(100); // o tamanho
            $objDrawing->setWorksheet($sheet); 

            $sheet->Cell('B2', 'Empresa : '.$empresa->nome);
            $sheet->Cell('B3', 'Contacto : '.$empresa->contacto_1);
            $sheet->Cell('B4', 'Nuit : '.$empresa->nuit);
            $sheet->Cell('B5', 'Enderenço :'.$empresa->morada);
            $sheet->Cell('B6', 'Email :'.$empresa->email);

            
            $sheet->cell('E2', function($cell) {

              // manipulate the cell
               $cell->setValue('Folha de Salário');
               $cell->setFont(array(
                 'family'     => 'Calibri',
                 'size'       => '16',
                 'bold' => true
                ));
            });
            $sheet->cell('D9', function($cell) {
               $cell->setValue('Remuneração');
               $cell->setFont(array(
                 'family'     => 'Calibri',
                 'size'       => '16',
                 'bold' => true
                ));
            });

            $sheet->cell('I9', function($cell) {
               $cell->setValue('Descontos');
               $cell->setFont(array(
                 'family'     => 'Calibri',
                 'size'       => '16',
                 'bold' => true
                ));
            });
            $sheet->cell('N9', function($cell) {
              $cell->setValue('Salários');
              $cell->setFont(array(
                'family'     => 'Calibri',
                'size'       => '16',
                'bold' => true
               ));
           });
          
            //method para formatar a celula C, usando casas decimais
            $sheet->setColumnFormat(array(
              'C10:C133' => '0.00',
             ));

             //dadosFolha => tras os dados que trazemos da base de dados.
             //A10 => define a posicacao onde ira cair o conteudo,
             //true =>permite que tenhas o cabecalho das colunas 
           $sheet->fromArray($dadosFolha, null, 'A10', true);
          
        });
        })->download('xls');
    }