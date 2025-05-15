install.packages("readxl")
install.packages("sem")
install.packages("lavaan")
install.packages("semTools")

library(readxl)
library(sem)
library(lavaan)
library(semTools)

psp_data <- read_excel("~/Documents/MSc 2022-23/Analysis 2/PSP3/PSPquantdata.xls")

# Extract numeric values from each column and combine into a single vector
#AMI
AMIc <- c(psp_data$Q24_1, psp_data$Q24_2, psp_data$Q24_3, psp_data$Q24_4, psp_data$Q24_5,  psp_data$Q24_6, psp_data$Q24_7, psp_data$Q24_8, psp_data$Q24_9, psp_data$Q24_10, psp_data$Q24_11, psp_data$Q24_12, psp_data$Q24_13, psp_data$Q24_14, psp_data$Q24_15, psp_data$Q24_16, psp_data$Q24_17, psp_data$Q24_18, psp_data$Q24_19, psp_data$Q24_20)
AMI <- as.numeric(AMIc)
summary(AMI, na.rm = TRUE)

#MSS
MSSc <- c(psp_data$Q25_1, psp_data$Q25_2, psp_data$Q25_3, psp_data$Q25_4, psp_data$Q25_5, psp_data$Q25_6, psp_data$Q25_7, psp_data$Q25_8)
MSS <- as.numeric(MSSc)
summary(MSS, na.rm = TRUE)

#BFI
BFIc <- c(psp_data$Q26_1, psp_data$Q26_2, psp_data$Q26_3, psp_data$Q26_4, psp_data$Q26_5, psp_data$Q26_6, psp_data$Q26_7, psp_data$Q26_8, psp_data$Q26_9, psp_data$Q26_10, psp_data$Q26_11, psp_data$Q26_12, psp_data$Q26_13, psp_data$Q26_14, psp_data$Q26_15, psp_data$Q26_16, psp_data$Q26_17, psp_data$Q26_18)
BFI <- as.numeric(BFIc)
summary(BFI, na.rm = TRUE)

#SDO
SDOc <- c(psp_data$Q27_1, psp_data$Q27_2, psp_data$Q27_3, psp_data$Q27_4, psp_data$Q27_5, psp_data$Q27_6, psp_data$Q27_7, psp_data$Q27_8, psp_data$Q27_9, psp_data$Q27_10, psp_data$Q27_11, psp_data$Q27_12, psp_data$Q27_13, psp_data$Q27_14, psp_data$Q27_15, psp_data$Q27_16)
SDO <- as.numeric(SDOc)
summary(SDO, na.rm = TRUE)

#RWA
RWAc <- c(psp_data$Q28_1, psp_data$Q28_2, psp_data$Q28_3, psp_data$Q28_4, psp_data$Q28_5, psp_data$Q28_6, psp_data$Q28_7, psp_data$Q28_8, psp_data$Q28_9, psp_data$Q28_10, psp_data$Q28_11, psp_data$Q28_12, psp_data$Q28_13, psp_data$Q28_14, psp_data$Q28_15)
RWA <- as.numeric(RWAc)
summary(RWA, na.rm = TRUE)

#Agreeableness
Agreeablenessc <- c(psp_data$Q26_1, psp_data$Q26_3, psp_data$Q26_5, psp_data$Q26_7, psp_data$Q26_9, psp_data$Q26_10, psp_data$Q26_12, psp_data$Q26_14, psp_data$Q26_17)
Agreeableness <- as.numeric(Agreeablenessc)
summary(Agreeableness, na.rm = TRUE)

#Openness
Opennessc <- c(psp_data$Q26_2, psp_data$Q26_4, psp_data$Q26_6, psp_data$Q26_8, psp_data$Q26_11, psp_data$Q26_13, psp_data$Q26_15, psp_data$Q26_16, psp_data$Q26_18)
Openness <- as.numeric(Opennessc)
summary(Openness, na.rm = TRUE)

#Gender
Genderc <- c(psp_data$Q17_1)
Gender <- as.numeric(Genderc)
summary(Gender, na.rm = TRUE)

#ASI 
ASIc <- c(psp_data$Q23_1, psp_data$Q23_2, psp_data$Q23_3, psp_data$Q23_4, psp_data$Q23_5, psp_data$Q23_6, psp_data$Q23_7, psp_data$Q23_8, psp_data$Q23_9, psp_data$Q23_10, psp_data$Q23_11, psp_data$Q23_12, psp_data$Q23_13, psp_data$Q23_14, psp_data$Q23_15, psp_data$Q23_16, psp_data$Q23_17, psp_data$Q23_18, psp_data$Q23_19, psp_data$Q23_20, psp_data$Q23_21, psp_data$Q23_22)
ASI <- as.numeric(ASIc)
summary(ASI, na.rm = TRUE)

maxlength <- max(length(AMI), length(BFI), length(MSS), length(ASI), length(Openness), length(Agreeableness), length(RWA), length(SDO), length(Gender))

ASIl = c(ASI, rep(NA, maxlength - length(ASI)))
AMIl = c(AMI, rep(NA, maxlength - length(AMI)))
MSSl = c(MSS, rep(NA, maxlength - length(MSS)))
Agreeablenessl = c(Agreeableness, rep(NA, maxlength - length(Agreeableness)))
Opennessl = c(Openness, rep(NA, maxlength - length(Openness)))
RWAl = c(RWA, rep(NA, maxlength - length(RWA)))
SDOl = c(SDO, rep(NA, maxlength - length(SDO)))
Genderl = c(Gender, rep(NA, maxlength - length(Gender)))

psp_frame <- data.frame(ASIl, AMIl, MSSl, Agreeablenessl, Opennessl, Genderl, RWAl, SDOl)


cor(psp_frame, use = "complete.obs", method = c("pearson"))



p_value_ASI_AMI <- cor.test(psp_frame$ASI, psp_frame$AMI)$p.value
p_value_ASI_MSS <- cor.test(psp_frame$ASI, psp_frame$MSS)$p.value
p_value_AMI_MSS <- cor.test(psp_frame$ASI, psp_frame$MSS)$p.value
p_value_AMI_MSS <- cor.test(psp_frame$AMI, psp_frame$MSS)$p.value
p_value_ASI_RWA <- cor.test(psp_frame$AMI, psp_frame$RWA)$p.value
p_value_ASI_RWA <- cor.test(psp_frame$ASI, psp_frame$RWA)$p.value
p_value_AMI_RWA <- cor.test(psp_frame$AMI, psp_frame$RWA)$p.value
p_value_AMI_SDO <- cor.test(psp_frame$AMI, psp_frame$SDO)$p.value
p_value_ASI_SDO <- cor.test(psp_frame$ASI, psp_frame$SDO)$p.value
p_value_ASI_Openness <- cor.test(psp_frame$Openness, psp_frame$ASI)$p.value
p_value_ASI_Agreeableness <- cor.test(psp_frame$Agreeableness, psp_frame$ASI)$p.value
p_value_AMI_Openness <- cor.test(psp_frame$Openness, psp_frame$AMI)$p.value
p_value_AMI_Agreeableness <- cor.test(psp_frame$Agreeableness, psp_frame$AMI)$p.value
p_value_MSS_Agreeableness <- cor.test(psp_frame$Agreeableness, psp_frame$MSS)$p.value
p_value_MSS_Openness <- cor.test(psp_frame$Openness, psp_frame$MSS)$p.value
p_value_Agreeableness_Openness <- cor.test(psp_frame$Openness, psp_frame$Agreeableness)$p.value
p_value_Agreeableness_SDO <- cor.test(psp_frame$SDO, psp_frame$Agreeableness)$p.value
p_value_Agreeableness_RWA <- cor.test(psp_frame$RWA, psp_frame$Agreeableness)$p.value
p_value_SDO_RWA <- cor.test(psp_frame$RWA, psp_frame$SDO)$p.value
p_value_SDO_Gender <- cor.test(psp_frame$RWA, psp_frame$Gender)$p.value
p_value_RWA_Gender <- cor.test(psp_frame$RWA, psp_frame$Gender)$p.value
p_value_Openness_Gender <- cor.test(psp_frame$Openness, psp_frame$Gender)$p.value
p_value_Agreeableness_Gender <- cor.test(psp_frame$Agreeableness, psp_frame$Gender)$p.value
p_value_ASI_Gender <- cor.test(psp_frame$ASI, psp_frame$Gender)$p.value
p_value_AMI_Gender <- cor.test(psp_frame$AMI, psp_frame$Gender)$p.value
p_value_MSS_Gender <- cor.test(psp_frame$MSS, psp_frame$Gender)$p.value
p_value_MSS_RWA <- cor.test(psp_frame$MSS, psp_frame$RWA)$p.value
p_value_Openness_SDO <- cor.test(psp_frame$Openness, psp_frame$SDO)$p.value
p_value_Openness_RWA <- cor.test(psp_frame$Openness, psp_frame$RWA)$p.value
print(p_value_ASI_AMI)
print(p_value_AMI_MSS)
print(p_value_ASI_RWA)
print(p_value_AMI_RWA)
print(p_value_AMI_SDO)
print(p_value_ASI_SDO)
print(p_value_ASI_Openness)
print(p_value_ASI_Agreeableness)
print(p_value_AMI_Openness)
print(p_value_AMI_Agreeableness)
print(p_value_MSS_Agreeableness)
print(p_value_MSS_Openness)
print(p_value_Agreeableness_Openness)
print(p_value_Agreeableness_SDO)
print(p_value_Agreeableness_RWA)
print(p_value_Openness_SDO)
print(p_value_Openness_RWA)
print(p_value_SDO_RWA)
print(p_value_SDO_Gender)
print(p_value_RWA_Gender)
print(p_value_Openness_Gender)
print(p_value_Agreeableness_Gender)
print(p_value_ASI_Gender)
print(p_value_AMI_Gender)
print(p_value_MSS_Gender)
print(p_value_MSS_RWA)

#Personality model - ASI
personality_model <- ' 
# Latent variable
    Sexism =~ RWAl + SDOl + Opennessl + Agreeablenessl + ASIl + Genderl
  # Residual variances of manifest variables
    RWAl ~~ e1*RWAl
    SDOl ~~ e2*SDOl
    Genderl ~~ e3*Genderl
    Opennessl ~~ e4*Opennessl
    Agreeablenessl ~~ e5*Agreeablenessl
    ASIl ~~ e6*ASIl 
    Opennessl ~~ e7*Agreeablenessl '
#paths?
#ASIl ~ SDOl 
#ASIl ~ RWAl 
#ASIl ~ Agreeablenessl 
#ASIl ~ Genderl 
#ASIl ~ Opennessl
#SDOl ~ Agreeablenessl 
#SDOl ~ Genderl 
#SDOl ~ RWAl 
#SDOl ~ Opennessl
#RWAl ~ Opennessl
#Agreeablenessl ~ Genderl '

fit <- sem(model = personality_model, data = psp_frame)
summary(fit, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

#Social psychology model - ASI
socialpsychology_model <- '
  # Latent variable
    Sexism =~ Genderl + Agreeablenessl + ASIl + Opennessl 
  # Residual variances of manifest variables
    RWAl ~~ e1*RWAl
    SDOl ~~ e2*SDOl
    Genderl ~~ e3*Genderl
    Opennessl ~~ e4*Opennessl
    Agreeablenessl ~~ e5*Agreeablenessl
    ASIl ~~ e6*ASIl 
    Agreeablenessl ~~ e7*Opennessl

    #paths?
Agreeablenessl ~ Genderl
ASIl ~ Genderl
SDOl ~ Agreeablenessl + Genderl +RWAl
RWAl ~ Opennessl

fit2 <- sem(model = socialpsychology_model, data = psp_frame)
summary(fit2, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

#Combined model - ASI
combined_model <- '
  # Latent variable
    Sexism =~ RWAl + SDOl + Genderl + Opennessl + Agreeablenessl + ASIl
  # Residual variances of manifest variables
    RWAl ~~ e1*RWAl
    SDOl ~~ e2*SDOl
    Genderl ~~ e3*Genderl
    Opennessl ~~ e4*Opennessl
    Agreeablenessl ~~ e5*Agreeablenessl
    ASIl ~~ e6*ASIl 
    #Paths?
    Agreeablenessl ~ Genderl
SDOl ~ Genderl
ASIl ~ Genderl
Sexism ~ Genderl
Sexism ~ SDOl
Sexism ~ RWAl
SDOl ~ RWAl
SDOl ~ Agreeablenessl
RWAl ~ Opennessl 
    
  
fit3 <- sem(model = combined_model, data = psp_frame)
summary(fit3, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

#Personality model - AMI
AMIpersonality_model <- ' 
# Latent variable
    Sexism =~ RWAl + SDOl + Opennessl + Agreeablenessl + AMIl
  # Residual variances of manifest variables
    RWAl ~~ e1*RWAl
    SDOl ~~ e2*SDOl
    Genderl ~~ e3*Genderl
    Opennessl ~~ e4*Opennessl
    Agreeablenessl ~~ e5*Agreeablenessl
    AMIl ~~ e6*AMIl 
    Agreeablenessl ~~ e7*Opennessl
#paths?
#Genderl ~ Agreeablenessl 
#AMIl ~ RWAl 
#AMIl ~ SDOl 
#SDOl ~ Agreeablenessl
#Genderl ~ SDOl + Agreeablenessl
#SDOl ~ RWAl
#RWAl ~ Opennessl '

fit4 <- sem(model = AMIpersonality_model, data = psp_frame)
summary(fit4, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

#Social psychology model - AMI
AMIsocialpsychology_model <- ' Agreeablenessl ~ Genderl
# Latent variable
Sexism =~ Genderl + Agreeablenessl + AMIl + Opennessl + SDOl + RWAl
# Residual variances of manifest variables
RWAl ~~ e1*RWAl
SDOl ~~ e2*SDOl
Genderl ~~ e3*Genderl
Opennessl ~~ e4*Opennessl
Agreeablenessl ~~ e5*Agreeablenessl
AMIl ~~ e6*AMIl
Agreeablenessl ~~ e7*Opennessl
 #paths?
    #Genderl ~ Agreeablenessl 
   # AMIl ~ Genderl
   # SDOl ~ Agreeablenessl
   # Genderl ~ SDOl + Agreeablenessl
   # SDOl ~ RWAl
  #  RWAl ~ Opennessl  

fit5 <- sem(model = AMIsocialpsychology_model, data = psp_frame)
summary(fit5, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

#Combined model - AMI
AMIcombined_model <- ' 
# Latent variable
Sexism =~ RWAl + SDOl + Genderl + Opennessl + Agreeablenessl + AMIl
# Residual variances of manifest variables
RWAl ~~ e1*RWAl
SDOl ~~ e2*SDOl
Genderl ~~ e3*Genderl
Opennessl ~~ e4*Opennessl
Agreeablenessl ~~ e5*Agreeablenessl
AMIl ~~ e6*AMIl 
Agreeablenessl ~~ e7*Opennessl 
#Paths?
Agreeablenessl ~ Genderl
SDOl ~ Genderl
AMIl ~ Genderl
Sexism ~ Genderl + SDOl + Agreeablenessl
Sexism ~ RWAl + Opennessl
SDOl ~ RWAl
SDOl ~ Agreeablenessl
RWAl ~ Opennessl 

fit6 <- sem(model = AMIcombined_model, data = psp_frame)
summary(fit6, standardized = TRUE, ci = TRUE, fit.measures = TRUE)

parTable(fit)
parTable(fit2)
parTable(fit3)
parTable(fit4)
parTable(fit5)
parTable(fit6)

install.packages("semTools")
library(semTools)

# Compare the fit of the models using various fit indices
#ASI only models
ASImodel_comparison1 <- semTools::compareFit(fit1, fit2, fit3)
print(model_comparison)
#AMI only models
AMImodel_comparison2 <- semTools::compareFit(fit4, fit5, fit6)
#All models
model_comparison3 <- semTools::compareFit(fit1, fit2, fit3, fit4, fit5, fit6)
#Combined models
model_combined <- semTools::compareFit(fit3, fit6)