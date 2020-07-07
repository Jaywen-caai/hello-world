# hello-world
test
Dim N(8), M(8), Pc(8), Tc(8), ωc(8) As Double:
Private Sub Command1_Click()


N(1) = Val(Text2.Text): N(2) = Val(Text8.Text): N(3) = Val(Text14.Text): N(4) = Val(Text20.Text): N(5) = Val(Text26.Text): N(6) = Val(Text32.Text): N(7) = Val(Text38.Text): N(8) = Val(Text44.Text):
M(1) = Val(Text3.Text): M(2) = Val(Text9.Text): M(3) = Val(Text15.Text): M(4) = Val(Text21.Text): M(5) = Val(Text27.Text): M(6) = Val(Text33.Text): M(7) = Val(Text39.Text): M(8) = Val(Text45.Text):
Pc(1) = Val(Text4.Text): Pc(2) = Val(Text10.Text): Pc(3) = Val(Text16.Text): Pc(4) = Val(Text22.Text): Pc(5) = Val(Text28.Text): Pc(6) = Val(Text34.Text): Pc(7) = Val(Text40.Text): Pc(8) = Val(Text46.Text):
Tc(1) = Val(Text5.Text): Tc(2) = Val(Text11.Text): Tc(3) = Val(Text17.Text): Tc(4) = Val(Text23.Text): Tc(5) = Val(Text29.Text): Tc(6) = Val(Text35.Text): Tc(7) = Val(Text41.Text): Tc(8) = Val(Text47.Text):
ωc(1) = Val(Text6.Text): ωc(2) = Val(Text12.Text): ωc(3) = Val(Text18.Text): ωc(4) = Val(Text24.Text): ωc(5) = Val(Text30.Text): ωc(6) = Val(Text36.Text): ωc(7) = Val(Text42.Text): ωc(8) = Val(Text48.Text):
P = Val(Text50.Text): T = Val(Text51.Text):

Ppc = 0: Tpc = 0: ωpc = 0: Mg = 0:
For i = 1 To 8
Ppc = Ppc + N(i) * Pc(i)
Tpc = Tpc + N(i) * Tc(i)
ωpc = ωpc + N(i) * ωc(i)
Mg = Mg + N(i) * M(i)
Next

R = 8.314
k = 0.48 + 1.574 * ωpc - 0.1715 * ωpc ^ 2
αt = (1 + k * (1 - (T / Tpc) ^ 0.5)) ^ 2
a = αt * 0.42742 * R ^ 2 * Tpc ^ 2 / Ppc
b = 0.08664 * R * Tpc / Ppc


'SRK状态方程直接迭代求得比容
V0 = R * Tpc / Ppc * 10 ^ (-6): Vk = V0:
Vk1 = V0 + b - (a * (Vk - b)) / (Ppc * Vk * (Vk + b))
While Abs(Vk1 - Vk) >= 0.0001
Vk = Vk1
Vk1 = V0 + b - (a * (Vk - b)) / (Ppc * Vk * (Vk + b))
Wend

'用比容求得页岩气质量密度
Text49.Text = Mg / Vk1


'用改进的SRK方程求比容（引入对比温度Tr,对比压力Pr)

Pr = P / Ppc: Tr = T / Tpc

αtr = 0.01543 * Tr ^ (-0.215) - 1

βpr = -0.00047 * Pr ^ 2 + 0.00158 * Pr - 0.0249

λ = 1 + a1 * αtr * βpr

V00 = R * Tr / Pr * 10 ^ (-6): Vk2 = V00:
Vk3 = (V00 + b - (a * (λ * Vk2 - b)) / (Pr * λ * Vk2 * (λ * Vk2 + b))) / λ
While Abs(Vk3 - Vk2) >= 0.0001
Vk2 = Vk3
Vk3 = (V00 + b - (a * (λ * Vk2 - b)) / (Pr * λ * Vk2 * (λ * Vk2 + b))) / λ
Wend
Text52.Text = Mg / Vk3


End Sub

Private Sub Command2_Click()
End
End Sub


