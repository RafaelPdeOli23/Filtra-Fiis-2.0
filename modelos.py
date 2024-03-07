class FundoImobiliario:
    def __init__(self, codigo, segmento, cotacao_atual, dividend_yield, p_pv, liquidez, vacancia_media):
        self.codigo = codigo
        self.segmento = segmento
        self.cotacao_atual = cotacao_atual
        self.dividend_yield = dividend_yield
        self.p_pv = p_pv
        self.liquidez = liquidez
        self.vacancia_media = vacancia_media


class Estrategia:
    def __init__(self, segmento ='', dividend_yield_minimo = 0, p_pv_maximo = 10,
                 liquidez_minima = 0, vacancia_media_maxima = 100):
        self.segmento = segmento
        self.dividend_yield_minimo = dividend_yield_minimo
        self.p_pv_maximo = p_pv_maximo
        self.liquidez_minima = liquidez_minima
        self.vacancia_media_maxima = vacancia_media_maxima

    def aplica_estrategia(self, fundo: FundoImobiliario):
        if self.segmento != "":
            if fundo.segmento != self.segmento:
                return False

        if fundo.dividend_yield < self.dividend_yield_minimo \
                or fundo.p_pv > self.p_pv_maximo \
                or fundo.liquidez < self.liquidez_minima\
                or fundo.vacancia_media > self.vacancia_media_maxima:
            return False
        else:
            return True
