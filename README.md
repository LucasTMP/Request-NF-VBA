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

A planilha foi desenvolvida para automatizar o processo de controle e registro dos usuários que estavam conectados em minha rede residencial (LAN e WLAN), ou seja a tabela DHCP, permitindo assim uma maior facilidade de acesso a essa informação. A programação foi construída para que a planilha conseguisse se conectar ao roteador/modem central, e capturar as informações necessárias, levando em conta que foi programada para o layout do roteador/modem fornecido pela empresa provedora de banda larga, já que é uma automação com foco em um problema local e singular. Para o uso em outros modelos basta alterar alguns parâmetros dentro do código. 

### Desenvolvido com

Esse projeto foi desenvolvido utilizando a ferramenta Excel do pacote Office, em conjunto com a programação em VBA (Visual Basic for Applications).

* [Excel](https://www.microsoft.com/pt-br/microsoft-365/excel)
* [Pacote Office](https://www.microsoft.com/pt-br/microsoft-365/business/compare-all-microsoft-365-business-products-b?&ef_id=EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE:G:s&OCID=AID2100139_SEM_EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE:G:s&lnkd=Google_O365SMB_Brand&gclid=EAIaIQobChMIw83Cx5HK8AIVwoORCh0CWwJ3EAAYASAAEgLcufD_BwE)


## Como começar

Basta realizar o download da planilha e abrir ela normalmente, aceitando os acessos e utilização de macros, **é necessário o conhecimento em programação para alterar o código de acordo com seu ambiente.**

### Pré-requisitos

Possuir a ferramenta Excel instalada, lembrando que a planilha está otimizada somente para o sistema operacional Windows (até o momento). Lembrando novamente que **é necessário o conhecimento em programação para alterar o código de acordo com seu ambiente.**


## Como usar

Primeiramente clique no botão *[ATUALIZAR]* para dar início a execução do programa, será aberta uma janela pedindo para que forneça a senha do roteador/modem e clique no botão *[Acessar]*, possibilitando assim o acesso ao sistema dele, após isso a planilha exibira uma outra janela indicando que a conexão com o aparelho está feita e mostrara o botão *[Pegar Dados]*, clicando nele a planilha processara as informações e exportara para o Excel. Com o termino do processamento será liberado o botão *[SAIR]* para fechar o programa.

Agora basta verificar os dados exportados, encontrando os aparelhos fisicamente que estão conectados na rede local pelo MAC ou IP dos aparelhos relatados, caso encontre realize o cadastro dele na planilha colocando o seu endereço MAC e a quem o aparelho pertence. Assim nas próximas vezes será possível usar os equipamentos já cadastrados para liberar o uso do campo de pesquisa, que faz uma busca do endereço MAC fornecido na exportação com os endereços MAC já cadastrados.


## Imagens


<p align="center" width="100%">
    <img width="33%" src="https://i.ibb.co/44mv7nW/nf2.jpg"> 
    <br>
    <em>Após clicar no botão [ATUALIZAR] a planilha exibe essa janela pedindo que o usuário forneça a senha de acesso ao aparelho (roteador/modem)</em>
</p>
 <hr>
 
<p align="center" width="100%">
    <img width="33%" src="https://i.ibb.co/zmbGnkc/mac3.jpg"> 
        <br>
    <em>Se a conexão for realizada com sucesso a aplicação exibira uma janela mostrando que está em contato com o aparelho e exibira o botão [Pegar Dados] para coletar as informações de DHCP.</em>
</p> 
 <hr>
 

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
