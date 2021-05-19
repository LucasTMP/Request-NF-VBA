[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

  <h1 align="center">Request-NF-VBA</h1>

<details open="open">
  <summary><b>Índice de Conteúdos</b></summary>
  <ol>
    <li>
      <a href="#Sobre-o-projeto">Sobre o Projeto</a>
      <ul>
        <li><a href="#Desenvolvido-com">Desenvolvido com</a></li>
      </ul>
    </li>
    <li>
      <a href="#Como-começar">Como começar</a>
      <ul>
        <li><a href="#Pré-requisitos ">Pré-requisitos </a></li>
      </ul>
    </li>
    <li><a href="#Como-usar">Como usar</a></li>
    <li><a href="#Imagens">Imagens</a></li>
    <li><a href="#Contato">Contato</a></li>
  </ol>
</details>

## Sobre o projeto

<p align="center" width="100%">
    <img width="33%" src="https://i.ibb.co/8mFLRKL/nf.png"> 
</p>

A planilha foi elaborada para sanar uma necessidade real de uma empresa do ramo industrial, que precisava automatizar seu processo de pedido de emissão de nota fiscal (NF) para o setor fiscal. Ela foi elaborada pensando em agregar praticidade para esse processo interno, visando realizar todo o procedimento em uma única aplicação, ou seja, preenchimento do formulário, transformação do formulário em PDF e envio do arquivo para o e-mail interno do setor fiscal seguindo um modelo de mensagem contendo informações sobre quem enviou e outros dados de rastreabilidade.

### Desenvolvido com

Esse projeto foi desenvolvido utilizando a ferramenta Excel do pacote Office, em conjunto com a programação em VBA (Visual Basic for Applications).

* [Excel](https://www.microsoft.com/pt-br/microsoft-365/excel)
* [Pacote Office](https://www.microsoft.com/pt-br/microsoft-365/business/compare-all-microsoft-365-business-products-b?&ef_id=EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE:G:s&OCID=AID2100139_SEM_EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE:G:s&lnkd=Google_O365SMB_Brand&gclid=EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE)


## Como começar

Basta realizar o download da planilha e abrir ela normalmente, aceitando os acessos e utilização de macros, **é necessário o conhecimento em programação para alterar o código de acordo com seu ambiente.**

### Pré-requisitos

Possuir a ferramenta Excel instalada, lembrando que a planilha está otimizada somente para o sistema operacional Windows (até o momento). Lembrando novamente que **é necessário o conhecimento em programação para alterar o código referente ao servidor SMTP que será utilizado, ao modelo de mensagem e ao local de criação dos arquivos (Excel,PDF).**


## Como usar

Para utilizar a planilha deve realizar o preenchimento dos campos do formulário de acordo com os dados solicitados, depois basta clicar no botão *[Enviar Formulário]* que a própria aplicação converterá o formulário criado em Excel para o PDF, salvando no local escolhido (definido como padrão em C:\ no código), estrutura um e-mail colocando o arquivo em anexo e realizando o envio da mensagem para o e-mail do setor financeiro.

Todos os parâmetros podem ser editados pelo VBA, como o servidor usado pelo SMTP, a conta utilizada para enviar o e-mail e o modelo de mensagem. Já os parâmetros do formulário estão configurados em planilhas ocultas, basta editá-los de acordo com a sua necessidade. 



## Imagens


<p align="center" width="100%">
    <img width="33%" src="https://i.ibb.co/44mv7nW/nf2.jpg"> 
    <br>
    <em>Trecho do código em VBA que precisa ser modificado para entram em acordo com as suas necessidades e ambiente de desenvolvimento e produção. </em>
</p>

## Contato

Lucas Teixeira - [Linkedin](https://www.linkedin.com/in/lucastmp/) 

<b>E-mail:</b> lucas.tmp@outlook.com


<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
<!-- MARKDOWN LINKS -->
[contributors-shield]: https://img.shields.io/github/contributors/LucasTMP/Request-NF-VBA.svg?style=for-the-badge
[contributors-url]: https://github.com/LucasTMP/Request-NF-VBA/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/LucasTMP/Request-NF-VBA.svg?style=for-the-badge
[forks-url]: https://github.com/LucasTMP/Request-NF-VBA/network/members
[stars-shield]: https://img.shields.io/github/stars/LucasTMP/Request-NF-VBA.svg?style=for-the-badge
[stars-url]: https://github.com/LucasTMP/Request-NF-VBA/stargazers
[issues-shield]: https://img.shields.io/github/issues/LucasTMP/Request-NF-VBA.svg?style=for-the-badge
[issues-url]: https://github.com/LucasTMP/Request-NF-VBA/issues
[license-shield]: https://img.shields.io/github/license/LucasTMP/Request-NF-VBA.svg?style=for-the-badge
[license-url]: https://github.com/LucasTMP/Request-NF-VBA/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://www.linkedin.com/in/lucastmp/
<!-- MARKDOWN IMAGES -->
