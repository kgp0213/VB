Attribute VB_Name = "modCrypt"

Public Function Crypt(texti, salasana) As String

       On Error Resume Next

              For T = 1 To Len(salasana)
                     sana = Asc(Mid(salasana, T, 1))
                     X1 = X1 + sana
              Next

       X1 = Int((X1 * 0.1) / 6)
       salasana = X1
       G = 0

              For TT = 1 To Len(texti)
                     sana = Asc(Mid(texti, TT, 1))
                     G = G + 1

                            If G = 6 Then G = 0
                                   X1 = 0

                                          If G = 0 Then X1 = sana - (salasana - 2)

                                                        If G = 1 Then X1 = sana + (salasana - 5)

                                                                      If G = 2 Then X1 = sana - (salasana - 4)

                                                                                    If G = 3 Then X1 = sana + (salasana - 2)

                                                                                                  If G = 4 Then X1 = sana - (salasana - 3)

                                                                                                                If G = 5 Then X1 = sana + (salasana - 5)
                                                                                                                       X1 = X1 + G
                                                                                                                       Crypted = Crypted & Chr(X1)
                                                                                                                Next


                                                                                                                Crypt = Crypted
                                                                                                                End Function


Public Function DeCrypt(texti, salasana) As String

       On Error Resume Next

              For T = 1 To Len(salasana)
                     sana = Asc(Mid(salasana, T, 1))
                     X1 = X1 + sana
              Next

       X1 = Int((X1 * 0.1) / 6)
       salasana = X1
       G = 0

              For TT = 1 To Len(texti)
                     sana = Asc(Mid(texti, TT, 1))
                     G = G + 1

                            If G = 6 Then G = 0
                                   X1 = 0

                                          If G = 0 Then X1 = sana + (salasana - 2)

                                                        If G = 1 Then X1 = sana - (salasana - 5)

                                                                      If G = 2 Then X1 = sana + (salasana - 4)

                                                                                    If G = 3 Then X1 = sana - (salasana - 2)

                                                                                                  If G = 4 Then X1 = sana + (salasana - 3)

                                                                                                                If G = 5 Then X1 = sana - (salasana - 5)
                                                                                                                       X1 = X1 - G
                                                                                                                       DeCrypted = DeCrypted & Chr(X1)
                                                                                                                Next


                                                                                                                DeCrypt = DeCrypted
                                                                                                                End Function
