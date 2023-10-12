# Controle-Estoque-VBA

O Controle de Estoque em Excel/VBA é um sistema robusto desenvolvido para atender às necessidades de pequenas empresas na gestão de estoque, cadastro de produtos e clientes. Este documento fornece informações sobre a instalação, utilização e suporte do sistema.

## Requisitos:

- Microsoft Excel 2016 ou superior;

## Instalação

1. Baixe o arquivo **Controle Estoque VBA.xlsm** do sistema no GitHub.
2. Abra o arquivo no Microsoft Excel.
3. Ao abrir o arquivo, clique em "Habilitar Conteúdo" para ativar as macros.

![habiliatr macro](https://media.discordapp.net/attachments/1109930711055618160/1162087069754085598/image.png?ex=653aa8eb&is=652833eb&hm=19fa2f344427fc88db2a9d7389f97de551cd876c0d2dcf860e10ea9628a3f46a&=&width=768&height=150)

# Utilização:

A seguir, apresentamos um guia passo a passo para utilizar o sistema, que oferece diversas funcionalidades:

![funcionalidades](https://media.discordapp.net/attachments/1109930711055618160/1162088610074464346/image.png?ex=653aaa5b&is=6528355b&hm=31c83a18e4558cbcfbd72cd4d7515e4433f8763cd7abd1ab54469391bdb5e0d2&=)

## Cadastrar Fornecedor ou Cliente

Ao clicar em "Cadastrar Fornecedor ou Cliente", abrirá uma caixa de diálogo na qual você deve preencher o campo com o nome do fornecedor ou cliente que deseja cadastrar.

![Cadastrar Fornecedor ou Cliente](https://media.discordapp.net/attachments/1109930711055618160/1162090056123699210/image.png?ex=653aabb3&is=652836b3&hm=226e19961d00cb3f7e5d2ea2887b892d0484c03f26c3b98c4873444b95f34161&=)

Clique em "Cadastrar" para confirmar o cadastro.

![mensagem](https://media.discordapp.net/attachments/1109930711055618160/1162090895588479046/image.png?ex=653aac7c&is=6528377c&hm=487e9ada53c8f52944576b15ff3133c603f8cae4731290e7e1be6e0c322f350d&=)

Os fornecedores, clientes e produtos ficam armazenados na aba "6-Cadastro Forn-Cli. e Prod."

## Cadastrar Produto

Ao clicar em "Cadastrar Produto", abrirá uma caixa de diálogo na qual você deve preencher os campos com os dados do produto.

![Cadastrar Produto](https://media.discordapp.net/attachments/1109930711055618160/1162161614523474043/image.png?ex=653aee58&is=65287958&hm=320089bf7a0fc6bc826b21a0aa44150a43088c9c99bae38de8385b0c7c0cc778&=)

Clique em "Cadastrar" para confirmar o cadastro.

![mensagem](https://media.discordapp.net/attachments/1109930711055618160/1162161887815938099/image.png?ex=653aee99&is=65287999&hm=67c2793bab886c7bfeadde2ec347c834f51fbe6f605909222142b864a4d61268&=)

Os fornecedores, clientes e produtos ficam armazenados na aba "6-Cadastro Forn-Cli. e Prod."

## Cadastrar Movimento

Ao clicar em "Cadastrar Movimento", abrirá uma caixa de diálogo na qual você deve preencher os campos com os dados do movimento, incluindo produtos e fornecedores ou clientes.

![Cadastrar Movimento](https://media.discordapp.net/attachments/1109930711055618160/1162163699885277294/image.png?ex=653af04a&is=65287b4a&hm=8194d0fa5e759290ad136dbd9048d0bf563c8e17c82d5de7ff5b788901d81804&=)

Produtos e Fornecedores ou Clientes são selecionados por meio de combobox, com base em cadastros anteriores.

![combobox](https://media.discordapp.net/attachments/1109930711055618160/1162164082728779858/image.png?ex=653af0a5&is=65287ba5&hm=26d7e4badbc4f44b53c94abc6dbae476bb18b9fb00fbaf36fc38f690fadb6aa2&=)

Clique em "Registrar" para confirmar o movimento.

![mensagem](https://media.discordapp.net/attachments/1109930711055618160/1162163781590331522/image.png?ex=653af05d&is=65287b5d&hm=758c340c285d2f5aa1dd69f7f1129775b369131fe704014b252ff1bb41449603&=)

A movimentação será armazenada na aba "3-Registro Entrada-Saida."

![movimentacao](https://media.discordapp.net/attachments/1109930711055618160/1162165358057226351/image.png?ex=653af1d5&is=65287cd5&hm=6f5667c640bac2e112866db8f0c036cfcc3f3896f4d5eeb34b80757ef57de599&=&width=768&height=80)

## Atualizar Estoque ou Atualizar Forn. e Clientes

Ao clicar em "Atualizar Estoque" ou "Atualizar Fornecedores e Clientes", os dados serão atualizados com as informações mais recentes.

![dados atualizado](https://media.discordapp.net/attachments/1109930711055618160/1162165809804750968/image.png?ex=653af241&is=65287d41&hm=a2fefbe8dd4b57f45e3e71f410d2b0c6d76e38f3fa4f53cbfa42bb10504e7b12&=)

A aba "4-Estoque" será atualizada com os estoques atualizados a partir das movimentações cadastradas.

![Estoque](https://media.discordapp.net/attachments/1109930711055618160/1162166342837882990/image.png?ex=653af2c0&is=65287dc0&hm=43f5d2fa7c8bf5e6419311179f8447cd0aa5eb18a3f1c15f1ee6fd6b18b8dca8&=&width=768&height=351)

A aba "5-FornXprodXqtd" será atualizada com informações sobre as relações comerciais com fornecedores ou clientes.

![controle fornecedores](https://media.discordapp.net/attachments/1109930711055618160/1162166661642731601/image.png?ex=653af30c&is=65287e0c&hm=c4030780a217d75a1c2bbea0871da52fa318db426dce716fc19b1b9dc4d7e7bd&=)

## Deletando Dados

Para excluir dados cadastrados, basta selecionar as células correspondentes e deletá-las. Após a exclusão, lembre-se de atualizar o estoque e os registros de fornecedores/clientes.

## Suporte e sugestões

Para obter suporte ou fornecer sugestões de melhoria para o sistema, entre em contato com o desenvolvedor pelo e-mail cardoso.henrique.93@gmail.com
