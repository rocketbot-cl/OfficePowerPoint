



# Office PowerPoint
  
Módulo para controlar o Microsoft PowerPoint.  

*Read this in other languages: [English](Manual_OfficePowerPoint.md), [Português](Manual_OfficePowerPoint.pr.md), [Español](Manual_OfficePowerPoint.es.md)*
  
![banner](imgs/Banner_OfficePowerPoint.png o jpg)
## Como instalar este módulo
  
Para instalar o módulo no Rocketbot Studio, pode ser feito de duas formas:
1. Manual: __Baixe__ o arquivo .zip e descompacte-o na pasta módulos. O nome da pasta deve ser o mesmo do módulo e dentro dela devem ter os seguintes arquivos e pastas: \__init__.py, package.json, docs, example e libs. Se você tiver o aplicativo aberto, atualize seu navegador para poder usar o novo módulo.
2. Automático: Ao entrar no Rocketbot Studio na margem direita você encontrará a seção **Addons**, selecione **Install Mods**, procure o módulo desejado e aperte instalar.  


## Descrição do comando

### Nova apresentação
  
Crie uma nova apresentação no PowerPoin
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |

### Abrir apresentação
  
Abre uma apresentação PowerPoint
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo||arquivo.pptx|

### Salvar apresentação
  
Salve a apresentação do PowerPoint que estava em edição
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Salvar arquivo||arquivo.pptx|

### Obter o tipo de slide
  
Lista todos os tipos de slides disponíveis
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado||res |

### Inserir slide
  
Insira um novo slide na apresentação
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Tipo|||

### Escrever dentro da apresentação
  
Escrever em uma apresentação do PowerPoint.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número do slide||0|
|ID ou nome do elemento onde será escrito||Title 1|
|Tamanho da fonte||8|
|Alinhar||Center|
|Negrito||True|
|Itálico||True|
|Sublinhar||True|
|Digitar texto||Lorem ipsum |

### Fechar a apresentação
  
Feche a apresentação que está sendo executada
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |

### Adicionar imagem
  
Adicione uma imagem à apresentação
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Rota da imagem||imagem.jpg|
|Número do slide||0|
|Posição||2,4.5 |
|Altura||5|

### Adicionar caixa de texto
  
Adicionar uma caixa de texto à apresentação.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Text||lorem ipsum|
|Posição||x,y|
|dimensão||Largura, Altura|
|Número do slide||0|

### Editar texto
  
Modifique o texto de um slide existente
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número do slide||0|
|Forma||Title 1|
|Tamanho da fonte||8|
|Alinhar||Center|
|Negrito||True|
|Itálico||True|
|Sublinhar||True|
|Texto||Lorem ipsum|

### Obter elementos da apresentação
  
Liste os elementos que estão na apresentação
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número do slide||0|
|Salvar o resultado||res|

### Editar as propriedades de um elemento
  
Edite as propriedades de um elemento em um slide
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número do slide||0|
|Id o nome do elemento||Title 1|
|Posição||x,y|
|Dimensão||Largura, Altura|
|Rotação||45.0|
