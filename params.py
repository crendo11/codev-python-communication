class Params:
    def __init__(self):
        
        # other properties
        self.lbda = 530e-09
        self.dI = 0.1

        
        # optical elements properties
        self.C0 = 0.007512
        self.f0 = 100e-03
        self.fe = 100e-03
        self.f1 = 0.075
        
        self.eta = 1.4607
        self.eta_air = 1
        

        # SLM properties
        self.SLMpitch = 8e-06
        self.slmWidth = 1920
        self.slmHeight = 1080

    def set_slmSize(self, width, height):
        self.slmWidth = width
        self.slmHeight = height