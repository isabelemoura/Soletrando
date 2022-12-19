dim p(20), n, pontos
dim i, nome, audio, acerto, errou

call carregar_voz

sub carregar_voz()
set audio=CreateObject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1

call inicio
end Sub

sub inicio()
nome=(inputbox("BEM VINDO AO SOLETRANDO!!!" + vbnewline &_
               "Supere seus limites!" + vbnewline &_
			   "***************************" + vbnewline &_
               "Digite o nome do Jogador", "Soletrando"))

             p(1)="country"
             p(2)="island"
			 p(3)="waffle"
			 p(4)="wafer"
			 p(5)="culture"
			 p(6)="cultural"
			 p(7)="chocolate"
			 p(8)="money"
			 p(9)="monkey"
			 p(10)="mickey"
			 p(11)="program"
             p(12)="apple"
             p(13)="to"
             p(14)="two"
             p(15)="three"
             p(16)="tree"
             p(17)="think"
             p(18)="thought"
             p(19)="sink"
             p(20)="answer"			 

call jogo
end Sub

sub jogo()
pontos = 0
acerto = 0
acertoN=0
i = 0
contO=0
contP=0
nivel=1

randomize(second(time))
n=int(rnd * 20) + 1
audio.speak (p(n))
do while i <= 20
	
	op=(inputbox("Jogador: "& nome &"" + vbnewline &_
	             "[O]Ouvir palavra novamente" + vbnewline &_
                 "[P]Pular palavra" + vbnewline &_			 
	             "Digite a palavra em letras min�sculas:", "Soletrando"))
				 
	select case op
	            case op=vbCancel
				     wscript.quit
	            case "p","P"
				    if contP=0 then
					   n=int(rnd * 20) + 1
					   audio.speak(p(n))
					   contP=contP+1
					   do while p(n) = "."
					       n=int(rnd * 20) + 1
						   audio.speak(p(n))
					    loop
					else 
					    msgbox("*Pular palavra* so poder ser usado apenas 1 vez no jogo!!"), vbInformation + vbOKOnly, "AVISO"
					end if
				case "o","O"
				    if contO=0 Then
					   audio.speak(p(n))
					   contO=contO+1
					else 
					    msgbox("*Ouvir palavra novamente* so pode ser usado apenas 1 vez no jogo!!"), vbInformation + vbOKOnly, "AVISO"
                    end if
				case else 
				    if op=p(n) Then
					       if acerto=9 Then
					          call acertou
							  
					       elseif acertoN=2 then
					       acerto=acerto+1
					       acertoN=0
						   acertoN=acertoN+1
						   nivel=nivel+1
					       pontos = pontos+(nivel*1000)
					       contV = contV + 1
					       msgBox("PARABENS "&nome&", VOCE ACERTOU!!!" + vbnewline &_
		                             "Palavra: "&p(n)&"" + vbnewline &_
                                     "Qnt de acertos: "& acerto &"" +vbnewline &_
                                     "Nivel "&nivel&"" + vbnewline &_
                                     "Pontuacao: "&pontos&"")
							    n=int(rnd * 20) + 1
						        audio.speak(p(n))
					        	do while p(n) = "."
						        n=int(rnd * 20) + 1
						    	audio.speak(p(n))
						        Loop
							
					     	elseif acertoN<2 then 
					    	acerto=acerto+1
					        acertoN=acertoN+1
					        pontos = pontos+(nivel*1000)
					        contV = contV + 1
					        msgBox("PARABENS "&nome&" VOCE ACERTOU!!!" + vbnewline &_
		                             "Palavra: "&p(n)&"" + vbnewline &_
                                     "Qnt de acertos: "& acerto &"" +vbnewline &_
                                     "Nivel "&nivel&"" + vbnewline &_
                                     "Pontuacao: "&pontos&"")
									 
							    n=int(rnd * 20) + 1
						        audio.speak(p(n))
					         	do while p(n) = "."
						        n=int(rnd * 20) + 1
						    	audio.speak(p(n))
						        Loop
					        end if
					else 
					    msgbox("Voce errou!!!" + vbnewline &_
						       "Qtde de Acertos: "& acerto &""), vbCritical + vbOKOnly, ""& nome &""
					    call jogarnvm
					end if
    end Select
Loop
end Sub

sub jogarnvm()
errou=msgbox("Deseja jogar novamente?",vbQuestion + vbYesNo, "AVISO")
if errou=vbyes Then
    call jogo
Else
    wscript.quit
end if
end Sub

sub acertou()
msgbox("Parabens voce VENCEU!!!"), vbInformation + vbOKOnly,"AVISO"
call jogarnvm
end sub

