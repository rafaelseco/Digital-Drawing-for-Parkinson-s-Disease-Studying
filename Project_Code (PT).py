import cv2 #importar a biblioteca cv2 -> webcam
import mediapipe as mp #importar a biblioteca do mediapipe
import time #importa a biblioteca referente ao tempo
import tkinter #importa a biblioteca para menus de dialogo com o utilizador
from tkinter import messagebox
import numpy as np
import inquirer
import math
import pandas as pd
from xlutils.copy import copy
import matplotlib.pyplot as plt
from matplotlib import style
import datetime as dt


def ratoclick(event, x, y, flags, param): 
  
    global Dificuldade,MenuInicial, pontosX, pontosY
    # checking for left mouse clicks 
    if event == cv2.EVENT_LBUTTONDOWN: 

        if (x>40 and x<200) and (y>375 and y<450):
            pontosX =  np.arange(0,440)
            pontosY =  100*np.sin(pontosX/28)+240
            pontosY = pontosY.astype(int)
            pontosX = pontosX + 100
            MenuInicial=1
            Dificuldade = 'Facil'

        if (x>240 and x<400) and (y>375 and y<450):
            pontosX =  np.arange(0,440)
            pontosY =  100*np.sin(1.4*pontosX/28)+240
            pontosY = pontosY.astype(int)
            pontosX = pontosX + 100
            MenuInicial=1
            Dificuldade = 'Medio'

        if (x>440 and x<600) and (y>375 and y<450):
            pontosX =  np.arange(0,440)
            pontosY =  100*np.sin(2*pontosX/28)+240
            pontosY = pontosY.astype(int)
            pontosX = pontosX + 100
            MenuInicial=1
            Dificuldade = 'Dificil'

    return Dificuldade, MenuInicial,pontosX,pontosY

if __name__=="__main__": 

    DATA = dt.datetime.today().strftime("%d/%m/%Y")
    tentativa = 0

    while True:
        ID = input("Insira o seu ID: ") #Para saber o ID de um paciente
        
        if ID.isdigit() and int(ID)>0:
            break
        else:
            print("Não introduziu um número válido!")


    FicheiroExcel_folha = pd.read_excel("Resultados.xlsx") #abre o Excel da pasta

    if len(FicheiroExcel_folha.index)>0:
        if int(ID) in FicheiroExcel_folha['ID'].values:
            for row in range(len(FicheiroExcel_folha.index)):
                if int(ID) == FicheiroExcel_folha['ID'][row]:
                    Novo_paciente = 0
                    NomePessoa = FicheiroExcel_folha['NOME'][row]
                    idade = FicheiroExcel_folha['IDADE'][row]
                    tentativa +=1
        else:
            Novo_paciente=1
    else:
        Novo_paciente=1
    
    if Novo_paciente==1:
        NomePessoa = input("Olá! Insira o seu nome: ") 
        tentativa = 0

        while True:
            idade = input("Insira a sua idade: ") 
            
            if idade.isdigit() and int(idade)>0:
                break
            else:
                print("Não introduziu um número!")
    else:
        print(f"\nOlá de novo, {NomePessoa}, espero que esteja tudo bem consigo!")


    #Para mostrar uma mensagem inicial:
    tkinter.Tk().geometry("200x200")
    avancar = messagebox.askokcancel(title="Informação Importante", message="O presente projeto está a ser desenvolvido por estudantes do Mestrado em Engenharia Biomédica no Instituto Superior de Engenharia do Porto, no âmbito da Unidade Curricular Interação Homem-Máquina. Ao continuar, está a permitir a abertura da câmara. Contudo, a sua imagem não será guardada.")  #retorna True se o botão OK for pressionado

    if avancar == True:

        mpMao = mp.solutions.hands #acede ao modulo sobre mãos presente na biblioteca do mediapipe
        maos = mpMao.Hands(max_num_hands=1) #tudo o que é preciso para agora trabalharmos com mãos

        mpDesenhar = mp.solutions.drawing_utils #esta é uma classe que ajuda a visualizar o resultado de uma tarefa do MediaPipe

        video = cv2.VideoCapture(0) #captura o video da webcam embutida no pc. Para uma webcam externa, index=1 (talvez)


        #Para calcular os FPS da camara:
        TempoInicial = 0
        
        #Temos de saber a primeira vez que o programa é corrido:
        global MenuInicial 
        MenuInicial = 0
        k=0
        CapturaVideo1=0
        CapturaVideo2=0
        CoordenadasDedoInicial=[0,0]
        mascara1 = 0
        mascara3 = 0
        mascara5 = 0
        mascara6 = 0
        desenhar = 0
        b=0
        a=0
        erroQuadraticoMedio = 0. #vai ser um float
        ciclo = 0
        erroQuadraticoMedioTotal = 0.
        Comeca_contar=0
        PercentagemAcerto=0
        segundos=0

        global Dificuldade, pontosX, pontosY
        Dificuldade = ''
        pontosX=[]
        pontosY=[]



        while (cv2.waitKey(1) != 27): #está sempre a rodar, apenas sai se a tecla pressionada for ESC. ESC corresponde a 27 na tabela ASCII
            
            abre, CapturaVideo = video.read() #lê o video. Retorna 1 (lógico) caso consiga abrir e 0 caso dê erro. Aloca o video/frames na variável 'CapturaVideo' para depois ser mostrado
            CapturaVideo = cv2.flip(CapturaVideo, 1) #faz o espelho do video
            videoLargura = video.get(cv2.CAP_PROP_FRAME_WIDTH) #qual a largura do frame da camara
            videoAltura = video.get(cv2.CAP_PROP_FRAME_HEIGHT) #qual a altura do frame da camara

            mascara = np.zeros_like(CapturaVideo) #a cada ciclo as mascaras serão uma imagem com tudo a zeros com as dimensões 640x480
            mascara2 = np.zeros_like(CapturaVideo)
            mascara4 = np.zeros_like(CapturaVideo)

            if not abre: #caso haja um erro na captura de video, faz um break ao ciclo while
                break


            TempoAgora=time.time() #vai buscar o tempo neste preciso instante, literalmente
            FPS = int(1/(TempoAgora-TempoInicial)) #a conversão para int deve-se a querermos numeros inteiros, senão teriamos numeros decimais
            TempoInicial = TempoAgora #é preciso atualizar a váriavel TempoInicial porque não pode ser sempre zero
            

            resultado = maos.process(cv2.cvtColor(CapturaVideo, cv2.COLOR_BGR2RGB)) #A ordem das cores do OpenCV é BGR, então para processar as imagens temos de converter para RGB para ter significado no futuro

            CoordenadasBotaoIniciar = [70,220] #Coordenadas de todos os botões presentes na imagem
            CoordenadasBotaoParar = [580,220]
            CoordenadasBotaoSair = [590,40]
            CoordenadasBotaoReset = [520,40]
            CoordenadasBotaoDificuldade = [450,40]

            CapturaVideo1 = CapturaVideo #A captura de video vai agora chamar-se CapturaVideo1 para podermos trabalhar sobre a imagem recebida e aplicar mascaras
            
            if MenuInicial==0:
                cv2.rectangle(CapturaVideo1,(40,375),(200,450),(0,255,0),-1)#cv2.rectangle(image, start_point, end_point, color, thickness)
                cv2.rectangle(CapturaVideo1,(240,375),(400,450),(0,255,255),-1)#cv2.rectangle(image, start_point, end_point, color, thickness)
                cv2.rectangle(CapturaVideo1,(440,375),(600,450),(0,0,255),-1)#cv2.rectangle(image, start_point, end_point, color, thickness)
                cv2.putText(CapturaVideo1, "FACIL",[85,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (255,255,255), 1)
                cv2.putText(CapturaVideo1, "MEDIO",[285,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (0,0,0), 1)
                cv2.putText(CapturaVideo1, "DIFICIL",[485,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (255,255,255), 1)
                cv2.putText(CapturaVideo1, "Ola, por favor escolha o nivel de dificuldade!",[100,60], cv2.FONT_HERSHEY_TRIPLEX, 0.5, (255,255,255), 1)
            
                cv2.namedWindow("Projeto de Interacao Homem-Maquina")
                cv2.setMouseCallback("Projeto de Interacao Homem-Maquina", ratoclick)
            
            else:

                if resultado.multi_hand_landmarks: #se conseguimos detetar uma mão

                    DedoIndicador = resultado.multi_hand_landmarks[0].landmark[mpMao.HandLandmark.INDEX_FINGER_TIP]  #vai focar no dedo indicador em vez de ser na mão toda
                    CoordenadasDedoFinal = mpDesenhar._normalized_to_pixel_coordinates(DedoIndicador.x, DedoIndicador.y, videoLargura, videoAltura) #vai calcular as coordenadas do dedo 
                        
                    if (CoordenadasDedoFinal!=None):
                    
                        cv2.circle(CapturaVideo1,CoordenadasBotaoIniciar,30,(0,255,0),-1)#parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.putText(CapturaVideo1, "Iniciar",[CoordenadasBotaoIniciar[0]-22,CoordenadasBotaoIniciar[1]+2], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.circle(CapturaVideo1,CoordenadasBotaoParar,30,(0,0,255),-1)#parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.putText(CapturaVideo1, "Parar",[CoordenadasBotaoParar[0]-20,CoordenadasBotaoParar[1]+5], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.circle(CapturaVideo1,CoordenadasBotaoSair,30,(0,0,0),-1)#parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.putText(CapturaVideo1, "X",[CoordenadasBotaoSair[0]-3,CoordenadasBotaoSair[1]+5], cv2.FONT_HERSHEY_TRIPLEX, 0.5, (255,255,255), 1)
                        cv2.circle(CapturaVideo1, CoordenadasBotaoReset, 30, (255,0,0), -1) #parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.putText(CapturaVideo1, "Limpar",[CoordenadasBotaoReset[0]-22,CoordenadasBotaoReset[1]+6], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (255,255,255), 1)
                        cv2.circle(CapturaVideo1, CoordenadasDedoFinal, 2, (76,171,206), 3) #parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.circle(CapturaVideo1, CoordenadasBotaoDificuldade, 30, (255,255,255), -1) #parâmetros de entrada: (image, center_coordinates, radius, color em BGR, thickness) ->  https://www.geeksforgeeks.org/python-opencv-cv2-circle-method/
                        cv2.putText(CapturaVideo1, "Mudar",[CoordenadasBotaoDificuldade[0]-22,CoordenadasBotaoDificuldade[1]+6], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.putText(CapturaVideo1, f"Dificuldade: {Dificuldade}",[10,15], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
        
                            #equação do Modulo da distancia entre 2 pontos:
                        equacaoIniciar = (CoordenadasBotaoIniciar[0]-CoordenadasDedoFinal[0])**2+(CoordenadasBotaoIniciar[1]-CoordenadasDedoFinal[1])**2
                        if (equacaoIniciar)<800:
                            desenhar = 1
                            tempo_inicio = time.time()
                            Comeca_contar = 1
                            

                        equacaoParar = (CoordenadasBotaoParar[0]-CoordenadasDedoFinal[0])**2+(CoordenadasBotaoParar[1]-CoordenadasDedoFinal[1])**2
                        if (equacaoParar)<800:
                            desenhar = 0
                            k = 0
                            if Comeca_contar == 1:
                                Comeca_contar = 0
                                tempo_final = time.time()
                                segundos = tempo_final - tempo_inicio
                                segundos = round(segundos,2)

                        equacaoSair = (CoordenadasBotaoSair[0]-CoordenadasDedoFinal[0])**2+(CoordenadasBotaoSair[1]-CoordenadasDedoFinal[1])**2
                        if equacaoSair<800:
                            if Comeca_contar == 0:
                                tempo_final = time.time()
                                segundos = tempo_final - tempo_inicio
                                segundos = round(segundos,2)
                            Sair =1
                            break

                        equacaoLimpar = (CoordenadasBotaoReset[0]-CoordenadasDedoFinal[0])**2+(CoordenadasBotaoReset[1]-CoordenadasDedoFinal[1])**2
                        if equacaoLimpar<800:
                            mascara1 = np.zeros_like(CapturaVideo)
                            mascara4 = np.zeros_like(CapturaVideo)
                            mascara2 = np.zeros_like(CapturaVideo)
                            mascara6 = 0

                            mascara3 = 0
                            mascara5 = 0
                            desenhar = 0
                            k = 0 
                            erroQuadraticoMedio = 0.
                            erroQuadraticoMedioTotal = 0.
                            ciclo = 0
                            Comeca_contar = 0
                            a=0

                        equacaoDiculdade = (CoordenadasBotaoDificuldade[0]-CoordenadasDedoFinal[0])**2+(CoordenadasBotaoDificuldade[1]-CoordenadasDedoFinal[1])**2
                        if (equacaoDiculdade)<800:
                            MenuInicial = 0
                            teste = np.zeros_like(CapturaVideo)
                            if ((mascara5|teste).any()!=0) and ciclo>0 and segundos!=0:
                                tentativa+=1
                                erroQuadraticoMedioTotal = erroQuadraticoMedioTotal/ciclo
                                erroQuadraticoMedioTotal = round(erroQuadraticoMedioTotal,2)
                                PercentagemAcerto = (np.sum(mascara6&mascara3)/np.sum(mascara3))*100                            
                                PercentagemAcerto = round(PercentagemAcerto,2)
                                NomeFicheiro = ID + '_' + NomePessoa + '_' + Dificuldade + '_'+ str(tentativa) + 'tentativa.jpg' #nome do ficheiro jpg
                                cv2.putText(mascara5, f"MSE: {str(erroQuadraticoMedioTotal)} ", (25,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Idade: {str(idade)} anos", (25,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Nome: {str(NomePessoa)} ", (220,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Dificuldade: {str(Dificuldade)} ", (400,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Desenho: ", (25,120), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Tempo: {str(segundos)} s", (350,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.putText(mascara5, f"Acerto: {str(PercentagemAcerto)} %", (350,430), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/                             
                                cv2.putText(mascara5, f"Data: {str(DATA)}", (350,460), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/                             
                                cv2.circle(mascara5,(50,425),5,(0,255,0),-1)
                                cv2.putText(mascara5, "MSE < 50", (57,428), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.circle(mascara5,(50,445),5,(0,255,255),-1)
                                cv2.putText(mascara5, "50 =< MSE < 70", (57,448), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                                cv2.circle(mascara5,(50,465),5,(0,0,255),-1)
                                cv2.putText(mascara5, "MSE >= 70", (57,468), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 

                                #FicheiroExcel_folha = pd.read_excel("Resultados.xlsx") #abre o Excel da pasta
                                novalinha = pd.DataFrame([[str(ID),NomePessoa,str(idade), Dificuldade,str(erroQuadraticoMedioTotal), str(PercentagemAcerto), str(segundos),str(DATA),str(tentativa)]], columns=['ID','NOME','IDADE','DIFICULDADE','ERRO','% ACERTO','TEMPO (s)','DATA','TENTATIVA'])
                                NovaTabela = pd.concat([FicheiroExcel_folha,novalinha])

                                NovaTabela.to_excel('Resultados.xlsx', index=False)

                                cv2.imwrite(NomeFicheiro, mascara5) #cria o ficheiro jpg -> https://www.geeksforgeeks.org/python-opencv-cv2-imwrite-method/
                            
                            mascara1 = np.zeros_like(CapturaVideo)
                            mascara4 = np.zeros_like(CapturaVideo)
                            mascara2 = np.zeros_like(CapturaVideo)

                            mascara3 = 0
                            mascara5 = 0
                            mascara6 = 0
                            desenhar = 0
                            k = 0 
                            erroQuadraticoMedio = 0.
                            erroQuadraticoMedioTotal = 0.
                            ciclo = 0
                            Comeca_contar = 0
                            a=0


                        if desenhar == 1:  
                            if k<1:
                                CoordenadasDedoInicial = CoordenadasDedoFinal

                                while a<439:  
                                    CoordenadasPontos = [pontosX[a],pontosY[a]]
                                    mascara2 = cv2.line(mascara2,CoordenadasPontos,[pontosX[a+1],pontosY[a+1]],(255,255,255),1)
                                    mascara3 = mascara3 + mascara2
                                    a=a+1

                                k+=1
                                
                            elif k==1:
                                i=0
                                erroQuadraticoMedio=0
                                pontos =  np.arange(40,610)
                                if CoordenadasDedoFinal[0] in pontos:
                                    mascara = cv2.line(mascara, CoordenadasDedoFinal, CoordenadasDedoInicial, (255,255,255), 5)  #para desenhar naquele ponto // parametros entrada: cv2.line(image, start_point, end_point, color, thickness)   -> https://www.geeksforgeeks.org/python-opencv-cv2-line-method/  
                                    
                                    if CoordenadasDedoFinal[0] in pontosX:

                                        i = list(pontosX).index(CoordenadasDedoFinal[0])
                                        erroQuadraticoMedio = math.fabs(CoordenadasDedoFinal[1]-pontosY[i])
                                        erroQuadraticoMedio = float(erroQuadraticoMedio**2)

                                        if erroQuadraticoMedio <50:
                                            mascara4 = cv2.line(mascara4, CoordenadasDedoFinal, CoordenadasDedoInicial, (0,255,0), 3)  #para desenhar naquele ponto // parametros entrada: cv2.line(image, start_point, end_point, color, thickness)   -> https://www.geeksforgeeks.org/python-opencv-cv2-line-method/  
                                        elif (erroQuadraticoMedio >=50 and erroQuadraticoMedio <70):
                                            mascara4 = cv2.line(mascara4, CoordenadasDedoFinal, CoordenadasDedoInicial, (0,255,255), 3)  #para desenhar naquele ponto // parametros entrada: cv2.line(image, start_point, end_point, color, thickness)   -> https://www.geeksforgeeks.org/python-opencv-cv2-line-method/  
                                        else:
                                            mascara4 = cv2.line(mascara4, CoordenadasDedoFinal, CoordenadasDedoInicial, (0,0,255), 3)  #para desenhar naquele ponto // parametros entrada: cv2.line(image, start_point, end_point, color, thickness)   -> https://www.geeksforgeeks.org/python-opencv-cv2-line-method/  

                                        mascara6 = mascara6 + mascara
                                        mascara5 = mascara5 + mascara4 + mascara3

                                    mascara1 = mascara1 + mascara + mascara3
                                    CapturaVideo1 = cv2.add(CapturaVideo1,mascara1)
                                    CoordenadasDedoInicial = CoordenadasDedoFinal

                                ciclo += 1
                                erroQuadraticoMedioTotal = float(erroQuadraticoMedioTotal + erroQuadraticoMedio)
                                
                    else:
                        print("O dedo não se encontra visível!")

                else:
                    cv2.putText(CapturaVideo1, "Nao consigo detetar a tua mao", (100,50), cv2.FONT_HERSHEY_TRIPLEX, 0.8, (0,0,0), 1)  #vai escrever  o numero de FPS no video // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 


            cv2.putText(CapturaVideo1, f"FPS: {str(FPS)}", (10,475), cv2.FONT_HERSHEY_TRIPLEX, 0.5, (76,171,206), 1)  #vai escrever  o numero de FPS no video // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
            cv2.imshow("Projeto de Interacao Homem-Maquina", CapturaVideo1) #mostra o video
        
        else:
            Sair = 0

        if Sair == 1:
            if ciclo>0:
                tentativa+=1
                erroQuadraticoMedioTotal = erroQuadraticoMedioTotal/ciclo
                erroQuadraticoMedioTotal = round(erroQuadraticoMedioTotal,2)
                PercentagemAcerto = (np.sum(mascara6&mascara3)/np.sum(mascara3))*100
                PercentagemAcerto = round(PercentagemAcerto,2)
                NomeFicheiro = ID + '_' + NomePessoa + '_' + Dificuldade + '_'+ str(tentativa) + 'tentativa.jpg' #nome do ficheiro jpg
                cv2.putText(mascara5, f"MSE: {str(erroQuadraticoMedioTotal)} ", (25,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Idade: {str(idade)} anos", (25,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Nome: {str(NomePessoa)} ", (220,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Dificuldade: {str(Dificuldade)} ", (400,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Desenho: ", (25,120), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Tempo: {str(segundos)} s", (350,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.putText(mascara5, f"Acerto: {str(PercentagemAcerto)} %", (350,430), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/                             
                cv2.putText(mascara5, f"Data: {str(DATA)}", (350,460), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/                                         
                cv2.circle(mascara5,(50,425),5,(0,255,0),-1)
                cv2.putText(mascara5, "MSE < 50", (57,428), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.circle(mascara5,(50,445),5,(0,255,255),-1)
                cv2.putText(mascara5, "50 =< MSE < 70", (57,448), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 
                cv2.circle(mascara5,(50,465),5,(0,0,255),-1)
                cv2.putText(mascara5, "MSE >= 70", (57,468), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  #vai escrever  o valor do erro Percentual quando o ficheiro for guardado // Sintaxe: cv2.putText(image, text, org, font, fontScale, color[, thickness[, lineType[, bottomLeftOrigin]]]) -> https://www.geeksforgeeks.org/python-opencv-cv2-puttext-method/ 


                #FicheiroExcel_folha = pd.read_excel("Resultados.xlsx") #abre o Excel da pasta

                novalinha = pd.DataFrame([[str(ID),NomePessoa,str(idade), Dificuldade,str(erroQuadraticoMedioTotal), str(PercentagemAcerto), str(segundos),str(DATA),str(tentativa)]], columns=['ID','NOME','IDADE','DIFICULDADE','ERRO','% ACERTO','TEMPO (s)','DATA','TENTATIVA'])
                NovaTabela = pd.concat([FicheiroExcel_folha,novalinha])

                NovaTabela.to_excel('Resultados.xlsx', index=False)

                cv2.imwrite(NomeFicheiro, mascara5) #cria o ficheiro jpg -> https://www.geeksforgeeks.org/python-opencv-cv2-imwrite-method/
            video.release() #Pára a captura de video
            cv2.destroyAllWindows() #fecha a janela

        elif Sair == 0:
            messagebox.showinfo("Atenção", "Premiu uma tecla ESC e, por isso, os seus dados não serão guardados!")

        else:
            messagebox.showinfo("Atenção", "Ocorreu um erro ao sair. Os seus dados não serão guardados.")

    if Novo_paciente==0 and Sair==1:
        tkinter.Tk().geometry("200x200")
        criargrafico = messagebox.askyesno(title="Informação Importante", message="Deseja criar um gráfico de follow-up com as avaliações anteriores?")  #retorna True se o botão OK for pressionado

        if criargrafico == True:
            ficheiro = pd.read_excel("Resultados.xlsx") #abre o Excel da pasta

            x=[]
            y=[]
            cores = []
            tamanho = []
            dificuldades = []

            for linha in range(len(ficheiro.index)):
                if int(ID) in ficheiro['ID'].values:
                    if int(ID) == ficheiro['ID'][linha]:
                        x.append(str(ficheiro['DATA'][linha]))
                        y.append(ficheiro['% ACERTO'][linha])
                        if int(ficheiro['ERRO'][linha])<250:
                            size = 1.3 * int(ficheiro['ERRO'][linha])
                            tamanho.append(size)
                        elif (int(ficheiro['ERRO'][linha])>=250) and (int(ficheiro['ERRO'][linha])<500):
                            size = 1.3 * int(ficheiro['ERRO'][linha])
                            tamanho.append(size)
                        else:
                            size = 1.2 * int(ficheiro['ERRO'][linha])
                            tamanho.append(size)
                        
                        if ficheiro['DIFICULDADE'][linha] == 'Facil':
                            cores.append('green')
                            dificuldades.append('Facil')
                        elif ficheiro['DIFICULDADE'][linha] == 'Medio':
                            cores.append('yellow')
                            dificuldades.append('Medio')
                        else:
                            cores.append('red')
                            dificuldades.append('Dificil')


            #plt.figure(figsize=(15,10))
            plt.style.use('seaborn-v0_8-dark-palette')
            #dificuldades = ['Facil','Medio','Dificil']
            
            fig, ax = plt.subplots()

            scatter = ax.scatter(x,y,marker="o",s=tamanho,edgecolors="black",c=cores)
            
            #legend1 = ax.legend(*scatter.legend_elements(),loc="upper right", title="Cores")
            #ax.add_artist(legend1)
            #ax.legend()
            ax.grid

            plt.xlabel("Data")
            plt.ylabel("% Acerto")
            plt.ylim(0,100)
            plt.title(f"Follow-up para o paciente com o ID {str(ID)} - Sr(a) {str(NomePessoa)} ({str(idade)} anos)")
            plt.show()