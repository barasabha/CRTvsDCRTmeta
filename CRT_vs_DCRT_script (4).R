########################################################################
### R-Script Publication                                             ###
###                                                                  ###
### Digital vs Conventional Cognitive Remediation Therapy (CRT) for  ###
### Improving Cognitive Performance in People with Mild Cognitive    ###
### Impairment (MCI) or Dementia                                     ### 
########################################################################

### Content
### 1) Loading packages
### 2) pairwise meta-analysis pre- to post- CRT (incl. sensitivity analyses)
### 3) pairwise meta-analysis pre- to post- DCRT (incl. sensitivity analyses)
### 4) pairwise meta-analysis CRT vs. ACC
### 5) pairwise meta-analysis DCRT vs. ACC (incl. sensitivity analysis)
### 6) network meta-analysis 
### 7) Network graph and crosstabulation of comparisons
### 8) Forest plot
### 9) Inconsistency checks 
###    9.1) The net heat plot
###    9.2) Net splitting
### 10) Comparison adjusted funnel plot


## set working directory
setwd("~/Desktop/CRT in MCI_Dementia/Analyses")


#####################################################################
### 1) Loading packages                                           ###
#####################################################################

#loading packages
library(readxl)
library(openxlsx)
library(metafor)
library(netmeta)

#####################################################################
### 2) pairwise meta-analysis pre- to post- CRT (incl. sensitivity analyses)###
#####################################################################

#data import
CRTprepost <- read_excel("CRT_pre_post.xlsx")

# Calculate effect sizes: Standardized Mean Change (within-subjects, pre-post)
CRT_pre_post <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,       
  sd1i = SD_MMSE_CRT_Outcome,        
  m2i = Mean_MMSE_CRT_Baseline,      
  sd2i = SD_MMSE_CRT_Baseline,       
  ni = nEG,                          
  ri = 0.5, # Assumed correlation between pre and post scores
  data = CRTprepost
)

# Run random-effects meta-analysis
rma_prepost_CRT <- rma(yi, vi, data = CRT_pre_post, method = "REML")

# Display summary of results
summary(rma_prepost_CRT)

# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -3.4345    6.8689   10.8689   10.0878   16.8689   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1600 (SE = 0.1468)
# tau (square root of estimated tau^2 value):      0.4000
# I^2 (total heterogeneity / total variability):   72.17%
# H^2 (total variability / sampling variability):  3.59
# 
# Test for Heterogeneity:
#   Q(df = 5) = 18.0143, p-val = 0.0029
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.4639  0.1974  2.3497  0.0188  0.0769  0.8508  * 

### outlier check pre-post CRT 
# Saving yi & vi in separate datasheet 
write.xlsx(CRT_pre_post, file = "CRT_pre_post_yi_vi.xlsx")

#outlier check
rstudent(rma_prepost_CRT)
#     resid     se       z 
# 1  0.0594 0.6031  0.0985 
# 2 -0.0128 0.5640 -0.0227 
# 3 -0.9474 0.4559 -2.0780 
# 4  0.0645 0.5818  0.1109 
# 5 -0.2537 0.6421 -0.3951 
# 6  0.7938 0.2602  3.0507 

influence(rma_prepost_CRT)
#   rstudent  dffits cook.d  cov.r tau2.del  QE.del    hat  weight    dfbs inf 
# 1   0.0985  0.0833 0.0084 1.4609   0.2128 17.9687 0.1534 15.3445  0.0824     
# 2  -0.0227  0.0344 0.0016 1.5999   0.2252 17.2790 0.2046 20.4557  0.0351     
# 3  -2.0780 -0.9441 0.6049 0.7111   0.0747  9.6162 0.1468 14.6811 -0.9907   * 
# 4   0.1109  0.0944 0.0113 1.5306   0.2194 17.9425 0.1776 17.7615  0.0948     
# 5  -0.3951 -0.1236 0.0167 1.2930   0.1907 17.3345 0.1176 11.7632 -0.1212     
# 6   3.0507  1.2531 0.4061 0.4436   0.0155  5.4187 0.1999 19.9941  1.1087   * 

#study 3 (De Luca et al., 2016) & study 6 (Tian et al., 2021) outlier

leave1out(rma_prepost_CRT)
#    estimate     se   zval   pval   ci.lb  ci.ub       Q     Qp   tau2      I2     H2 
# -1   0.4458 0.2386 1.8683 0.0617 -0.0219 0.9135 17.9687 0.0013 0.2128 78.5890 4.6705 
# -2   0.4560 0.2497 1.8262 0.0678 -0.0334 0.9454 17.2790 0.0017 0.2252 74.7812 3.9653 
# -3   0.6174 0.1665 3.7086 0.0002  0.2911 0.9437  9.6162 0.0474 0.0747 56.7241 2.3108 
# -4   0.4429 0.2442 1.8132 0.0698 -0.0358 0.9216 17.9425 0.0013 0.2194 77.5526 4.4549 
# -5   0.4894 0.2245 2.1799 0.0293  0.0494 0.9293 17.3345 0.0017 0.1907 77.9534 4.5358 
# -6   0.3381 0.1315 2.5711 0.0101  0.0804 0.5958  5.4187 0.2470 0.0155 17.4318 1.2111 

# Create forest plot
jpeg("Forest Plot - pre-post CRT.jpeg", width=600, height=400)
forest(rma_prepost_CRT,
       slab = paste(CRT_pre_post$Study.name, sep=""),
       psize = 1, lwd = 2,
       cex = .90,
       header = c("Reference", "Weight (%) SMDs [95% CI]"),
       addpred = TRUE,
       predstyle = "dist",
       showweights = TRUE)
usr <- par("usr")
x_pos <- usr[1] 
text(x_pos, -1.5, pos=4, cex=0.9, 
     bquote(paste("(Q = ", .(formatC(rma_prepost_CRT$QE, digits=2, format="f")), 
                  ", df = ", .(rma_prepost_CRT$k - rma_prepost_CRT$p), 
                  ", p = ", .(formatC(rma_prepost_CRT$QEp, digits=2, format="f")), 
                  "; ", I^2, " = ", .(formatC(rma_prepost_CRT$I2, digits=1, format="f")), "%)")))
dev.off() 


#funnel plot
jpeg("funnel_plot_CRT_prepost.jpeg")
funnel(rma_prepost_CRT)
dev.off()

#Egger´s test
#k < 10

######################## sensitivity analyses ##################################
# Test model robustness across different r assumptions

### sensitivity analysis a) r = 0.2
CRT_pre_post0.2 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,       
  sd1i = SD_MMSE_CRT_Outcome,        
  m2i = Mean_MMSE_CRT_Baseline,      
  sd2i = SD_MMSE_CRT_Baseline,       
  ni = nEG,                          
  ri = 0.2, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.2 <- rma(yi, vi, data = CRT_pre_post0.2, method = "REML")
summary (rma_prepost_CRT0.2)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -2.5390    5.0781    9.0781    8.2969   15.0781   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0970 (SE = 0.1031)
# tau (square root of estimated tau^2 value):      0.3114
# I^2 (total heterogeneity / total variability):   62.45%
# H^2 (total variability / sampling variability):  2.66
# 
# Test for Heterogeneity:
#   Q(df = 5) = 13.5171, p-val = 0.0190
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.3920  0.1659  2.3627  0.0181  0.0668  0.7172  * 

### sensitivity analysis b) r = 0.3
CRT_pre_post0.3 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,       
  sd1i = SD_MMSE_CRT_Outcome,        
  m2i = Mean_MMSE_CRT_Baseline,      
  sd2i = SD_MMSE_CRT_Baseline,       
  ni = nEG,                          
  ri = 0.3, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.3 <- rma(yi, vi, data = CRT_pre_post0.3, method = "REML")
summary (rma_prepost_CRT0.3)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -2.7972    5.5944    9.5944    8.8133   15.5944   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1128 (SE = 0.1142)
# tau (square root of estimated tau^2 value):      0.3358
# I^2 (total heterogeneity / total variability):   65.58%
# H^2 (total variability / sampling variability):  2.91
# 
# Test for Heterogeneity:
#   Q(df = 5) = 14.7395, p-val = 0.0115
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.4109  0.1744  2.3557  0.0185  0.0690  0.7527  *

### sensitivity analysis c) r = 0.4
CRT_pre_post0.4 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,       
  sd1i = SD_MMSE_CRT_Outcome,        
  m2i = Mean_MMSE_CRT_Baseline,      
  sd2i = SD_MMSE_CRT_Baseline,       
  ni = nEG,                          
  ri = 0.4, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.4 <- rma(yi, vi, data = CRT_pre_post0.4, method = "REML")
summary (rma_prepost_CRT0.4)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -3.0920    6.1840   10.1840    9.4029   16.1840   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1330 (SE = 0.1283)
# tau (square root of estimated tau^2 value):      0.3647
# I^2 (total heterogeneity / total variability):   68.82%
# H^2 (total variability / sampling variability):  3.21
# 
# Test for Heterogeneity:
#   Q(df = 5) = 16.2096, p-val = 0.0063
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.4341  0.1847  2.3508  0.0187  0.0722  0.7961  *


### sensitivity analysis d) r = 0.6
CRT_pre_post0.6 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,       
  sd1i = SD_MMSE_CRT_Outcome,        
  m2i = Mean_MMSE_CRT_Baseline,      
  sd2i = SD_MMSE_CRT_Baseline,       
  ni = nEG,                          
  ri = 0.6, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.6 <- rma(yi, vi, data = CRT_pre_post0.6, method = "REML")
summary (rma_prepost_CRT0.6)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -3.8419    7.6839   11.6839   10.9027   17.6839   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1976 (SE = 0.1725)
# tau (square root of estimated tau^2 value):      0.4446
# I^2 (total heterogeneity / total variability):   75.63%
# H^2 (total variability / sampling variability):  4.10
# 
# Test for Heterogeneity:
#   Q(df = 5) = 20.2902, p-val = 0.0011
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.5037  0.2138  2.3560  0.0185  0.0847  0.9227  *


### sensitivity analysis e) r = 0.7
CRT_pre_post0.7 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,
  sd1i = SD_MMSE_CRT_Outcome,
  m2i = Mean_MMSE_CRT_Baseline,
  sd2i = SD_MMSE_CRT_Baseline,
  ni = nEG,
  ri = 0.7, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.7 <- rma(yi, vi, data = CRT_pre_post0.7, method = "REML")
summary (rma_prepost_CRT0.7)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -4.3441    8.6881   12.6881   11.9070   18.6881   
# 
# tau^2 (estimated amount of total heterogeneity): 0.2536 (SE = 0.2105)
# tau (square root of estimated tau^2 value):      0.5036
# I^2 (total heterogeneity / total variability):   79.18%
# H^2 (total variability / sampling variability):  4.80
# 
# Test for Heterogeneity:
#   Q(df = 5) = 23.2723, p-val = 0.0003
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.5609  0.2359  2.3772  0.0174  0.0985  1.0233  * 


### sensitivity analysis f) r = 0.8
CRT_pre_post0.8 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_CRT_Outcome,
  sd1i = SD_MMSE_CRT_Outcome,
  m2i = Mean_MMSE_CRT_Baseline,
  sd2i = SD_MMSE_CRT_Baseline,
  ni = nEG,
  ri = 0.8, # Assumed correlation between pre and post scores
  data = CRTprepost
)
rma_prepost_CRT0.8 <- rma(yi, vi, data = CRT_pre_post0.8, method = "REML")
summary (rma_prepost_CRT0.8)
# Random-Effects Model (k = 6; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -5.0005   10.0010   14.0010   13.2198   20.0010   
# 
# tau^2 (estimated amount of total heterogeneity): 0.3466 (SE = 0.2732)
# tau (square root of estimated tau^2 value):      0.5888
# I^2 (total heterogeneity / total variability):   82.82%
# H^2 (total variability / sampling variability):  5.82
# 
# Test for Heterogeneity:
#   Q(df = 5) = 27.4385, p-val < .0001
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub    
# 0.6536  0.2686  2.4328  0.0150  0.1270  1.1801  * 


### significant change in MMSE scores from pre- to post in the CRT group.




#####################################################################
### 3) pairwise meta-analysis pre- to post- DCRT                  ###
#####################################################################

#data import
DCRTprepost <- read_excel("DCRT_pre_post.xlsx")

# Calculate effect sizes: Standardized Mean Change (within-subjects, pre-post)
DCRT_pre_post <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.5, # Assumed correlation between pre and post scores
  data = DCRTprepost
)

# Run random-effects meta-analysis
rma_prepost_DCRT <- rma(yi, vi, data = DCRT_pre_post, method = "REML")

# Display summary of results
summary(rma_prepost_DCRT)

# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -6.3665   12.7330   16.7330   17.7028   18.0663   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0614 (SE = 0.0528)
# tau (square root of estimated tau^2 value):      0.2479
# I^2 (total heterogeneity / total variability):   51.46%
# H^2 (total variability / sampling variability):  2.06
# 
# Test for Heterogeneity:
#   Q(df = 12) = 23.2590, p-val = 0.0256
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.2802  0.1019  2.7511  0.0059  0.0806  0.4799  ** 

### outlier check pre-post DCRT 
# Saving yi & vi in separate datasheet 
write.xlsx(DCRT_pre_post, file = "DCRT_pre_post_yi_vi.xlsx")
#outlier check
rstudent(rma_prepost_DCRT)
#      resid     se       z 
# 1   0.1758 0.4306  0.4082 
# 2   0.4317 0.4007  1.0774 
# 3   0.6554 0.4407  1.4871 
# 4  -0.6302 0.3476 -1.8131 
# 5  -0.0217 0.3569 -0.0607 
# 6   0.2317 0.4916  0.4712 
# 7   0.2157 0.4413  0.4888 
# 8  -0.1523 0.3380 -0.4507 
# 9  -0.8563 0.3133 -2.7336 
# 10 -0.1358 0.4371 -0.3107 
# 11  0.8049 0.4457  1.8060 
# 12  0.0057 0.3789  0.0150 
# 13  0.0216 0.3833  0.0564 

influence(rma_prepost_DCRT)
#    rstudent  dffits cook.d  cov.r tau2.del  QE.del    hat  weight    dfbs inf 
# 1    0.4082  0.0748 0.0060 1.1746   0.0732 22.8022 0.0643  6.4260  0.0739     
# 2    1.0774  0.3051 0.0894 1.0145   0.0553 21.0983 0.0664  6.6421  0.3071     
# 3    1.4871  0.3883 0.1382 0.8996   0.0446 20.0711 0.0514  5.1436  0.4017     
# 4   -1.8131 -0.3933 0.1270 0.8517   0.0370 18.9530 0.0761  7.6085 -0.4010     
# 5   -0.0607 -0.0876 0.0101 1.3798   0.0890 23.2139 0.1213 12.1315 -0.0910     
# 6    0.4712  0.0836 0.0072 1.1174   0.0691 22.8131 0.0467  4.6660  0.0825     
# 7    0.4888  0.0974 0.0100 1.1498   0.0710 22.6927 0.0599  5.9898  0.0963     
# 8   -0.4507 -0.2241 0.0674 1.3820   0.0869 22.4582 0.1394 13.9449 -0.2362     
# 9   -2.7336 -0.2701 0.0468 0.5543   0.0100 15.2332 0.0721  7.2143 -0.2996     
# 10  -0.3107 -0.1286 0.0182 1.2098   0.0773 23.1922 0.0638  6.3781 -0.1267     
# 11   1.8060  0.4658 0.1915 0.8153   0.0360 18.9745 0.0481  4.8132  0.4952     
# 12   0.0150 -0.0551 0.0037 1.3143   0.0846 23.1900 0.0972  9.7184 -0.0558     
# 13   0.0564 -0.0402 0.0019 1.3008   0.0836 23.1592 0.0932  9.3237 -0.0405 

leave1out(rma_prepost_DCRT)
#     estimate     se   zval   pval  ci.lb  ci.ub       Q     Qp   tau2      I2     H2 
# -1    0.2723 0.1104 2.4670 0.0136 0.0560 0.4887 22.8022 0.0188 0.0732 56.5956 2.3039 
# -2    0.2498 0.1026 2.4344 0.0149 0.0487 0.4509 21.0983 0.0324 0.0553 49.5169 1.9809 
# -3    0.2424 0.0966 2.5086 0.0121 0.0530 0.4317 20.0711 0.0444 0.0446 44.6481 1.8066 
# -4    0.3165 0.0940 3.3670 0.0008 0.1323 0.5008 18.9530 0.0619 0.0370 39.2907 1.6472 
# -5    0.2905 0.1197 2.4278 0.0152 0.0560 0.5250 23.2139 0.0165 0.0890 57.3439 2.3443 
# -6    0.2716 0.1077 2.5221 0.0117 0.0605 0.4826 22.8131 0.0188 0.0691 55.6942 2.2570 
# -7    0.2700 0.1092 2.4723 0.0134 0.0560 0.4841 22.6927 0.0195 0.0710 55.9883 2.2721 
# -8    0.3067 0.1197 2.5611 0.0104 0.0720 0.5414 22.4582 0.0211 0.0869 53.8721 2.1679 
# -9    0.3023 0.0758 3.9858 0.0001 0.1536 0.4509 15.2332 0.1721 0.0100 14.9551 1.1758 
# -10   0.2940 0.1120 2.6237 0.0087 0.0744 0.5136 23.1922 0.0166 0.0773 57.9264 2.3768 
# -11   0.2357 0.0920 2.5621 0.0104 0.0554 0.4159 18.9745 0.0616 0.0360 39.5567 1.6544 
# -12   0.2864 0.1168 2.4527 0.0142 0.0575 0.5153 23.1900 0.0166 0.0846 58.5098 2.4102 
# -13   0.2847 0.1162 2.4508 0.0143 0.0570 0.5124 23.1592 0.0168 0.0836 58.4764 2.4083 

## Lu et al., 2024 outlier (?)

# Create forest plot
jpeg("Forest Plot - pre-post DCRT.jpeg", width=600, height=700)
forest(rma_prepost_DCRT,
       slab = paste(DCRT_pre_post$Study.name, sep=""),
       psize = 1, lwd = 2,
       cex = .90,
       header = c("Reference", "Weight (%) SMDs [95% CI]"),
       addpred = TRUE,
       predstyle = "dist",
       showweights = TRUE)
usr <- par("usr")
x_pos <- usr[1] 
text(x_pos, -1.5, pos=4, cex=0.9, 
     bquote(paste("(Q = ", .(formatC(rma_prepost_DCRT$QE, digits=2, format="f")), 
                  ", df = ", .(rma_prepost_DCRT$k - rma_prepost_CRT$p), 
                  ", p = ", .(formatC(rma_prepost_DCRT$QEp, digits=2, format="f")), 
                  "; ", I^2, " = ", .(formatC(rma_prepost_DCRT$I2, digits=1, format="f")), "%)")))
dev.off() 

#funnel plot
jpeg("funnel_plot_DCRT_prepost.jpeg")
funnel (rma_prepost_DCRT)
dev.off()

#Egger´s test
regtest(rma_prepost_DCRT)
# Regression Test for Funnel Plot Asymmetry
# 
# Model:     mixed-effects meta-regression model
# Predictor: standard error
# 
# Test for Funnel Plot Asymmetry: z =  1.3673, p = 0.1715
# Limit Estimate (as sei -> 0):   b = -0.1276 (CI: -0.7496, 0.4943)


######################## sensitivity analyses #################################
# Test model robustness across different r assumptions
### sensitivity analysis a) r = 0.2
DCRT_pre_post0.2 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.2, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.2 <- rma(yi, vi, data = DCRT_pre_post0.2, method = "REML")
summary(rma_prepost_DCRT0.2)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -3.1223    6.2446   10.2446   11.2145   11.5780   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0000 (SE = 0.0175)
# tau (square root of estimated tau^2 value):      0.0004
# I^2 (total heterogeneity / total variability):   0.00%
# H^2 (total variability / sampling variability):  1.00
# 
# Test for Heterogeneity:
#   Q(df = 12) = 15.8229, p-val = 0.1995
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.1976  0.0631  3.1309  0.0017  0.0739  0.3212  ** 


### sensitivity analysis b) r = 0.3
DCRT_pre_post0.3 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.3, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.3 <- rma(yi, vi, data = DCRT_pre_post0.3, method = "REML")
summary(rma_prepost_DCRT0.3)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -4.1181    8.2362   12.2362   13.2060   13.5696   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0063 (SE = 0.0222)
# tau (square root of estimated tau^2 value):      0.0793
# I^2 (total heterogeneity / total variability):   9.97%
# H^2 (total variability / sampling variability):  1.11
# 
# Test for Heterogeneity:
#   Q(df = 12) = 17.7027, p-val = 0.1250
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.2168  0.0698  3.1073  0.0019  0.0801  0.3536  ** 


### sensitivity analysis c) r = 0.4
DCRT_pre_post0.4 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.4, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.4 <- rma(yi, vi, data = DCRT_pre_post0.4, method = "REML")
summary(rma_prepost_DCRT0.4)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -5.1904   10.3808   14.3808   15.3506   15.7142   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0321 (SE = 0.0378)
# tau (square root of estimated tau^2 value):      0.1792
# I^2 (total heterogeneity / total variability):   35.92%
# H^2 (total variability / sampling variability):  1.56
# 
# Test for Heterogeneity:
#   Q(df = 12) = 20.0978, p-val = 0.0653
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.2492  0.0875  2.8484  0.0044  0.0777  0.4206  ** 


### sensitivity analysis d) r = 0.6
DCRT_pre_post0.6 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.6, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.6 <- rma(yi, vi, data = DCRT_pre_post0.6, method = "REML")
summary(rma_prepost_DCRT0.6)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -7.7376   15.4752   19.4752   20.4451   20.8086   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1016 (SE = 0.0719)
# tau (square root of estimated tau^2 value):      0.3187
# I^2 (total heterogeneity / total variability):   63.29%
# H^2 (total variability / sampling variability):  2.72
# 
# Test for Heterogeneity:
#   Q(df = 12) = 27.6361, p-val = 0.0063
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.3178  0.1179  2.6956  0.0070  0.0867  0.5489  ** 

### sensitivity analysis d) r = 0.6
DCRT_pre_post0.6 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.6, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.6 <- rma(yi, vi, data = DCRT_pre_post0.6, method = "REML")
summary(rma_prepost_DCRT0.6)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -7.7376   15.4752   19.4752   20.4451   20.8086   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1016 (SE = 0.0719)
# tau (square root of estimated tau^2 value):      0.3187
# I^2 (total heterogeneity / total variability):   63.29%
# H^2 (total variability / sampling variability):  2.72
# 
# Test for Heterogeneity:
#   Q(df = 12) = 27.6361, p-val = 0.0063
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.3178  0.1179  2.6956  0.0070  0.0867  0.5489  ** 


### sensitivity analysis e) r = 0.7
DCRT_pre_post0.7 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.7, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.7 <- rma(yi, vi, data = DCRT_pre_post0.7, method = "REML")
summary(rma_prepost_DCRT0.7)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -9.4341   18.8682   22.8682   23.8380   24.2015   
# 
# tau^2 (estimated amount of total heterogeneity): 0.1644 (SE = 0.1006)
# tau (square root of estimated tau^2 value):      0.4055
# I^2 (total heterogeneity / total variability):   73.13%
# H^2 (total variability / sampling variability):  3.72
# 
# Test for Heterogeneity:
#   Q(df = 12) = 34.1333, p-val = 0.0006
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.3684  0.1387  2.6551  0.0079  0.0964  0.6403  ** 


### sensitivity analysis f) r = 0.8
DCRT_pre_post0.8 <- escalc(
  measure = "SMCC",
  m1i = Mean_MMSE_DCRT_Outcome,       
  sd1i = SD_MMSE_DCRT_Outcome,        
  m2i = Mean_MMSE_DCRT_Baseline,      
  sd2i = SD_MMSE_DCRT_Baseline,       
  ni = nEG,                          
  ri = 0.8, # Assumed correlation between pre and post scores
  data = DCRTprepost
)
rma_prepost_DCRT0.8 <- rma(yi, vi, data = DCRT_pre_post0.8, method = "REML")
summary(rma_prepost_DCRT0.8)
# Random-Effects Model (k = 13; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -11.7219   23.4438   27.4438   28.4136   28.7772   
# 
# tau^2 (estimated amount of total heterogeneity): 0.2820 (SE = 0.1529)
# tau (square root of estimated tau^2 value):      0.5311
# I^2 (total heterogeneity / total variability):   81.73%
# H^2 (total variability / sampling variability):  5.47
# 
# Test for Heterogeneity:
#   Q(df = 12) = 44.9190, p-val < .0001
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.4455  0.1705  2.6126  0.0090  0.1113  0.7798  ** 


### significant change in MMSE scores from pre- to post in the DCRT group.



#####################################################################
### 4) pairwise meta-analysis CRT vs. ACC.                        ###
#####################################################################

#data import
CRTvsACC <- read_excel("CRT_vs_ACC.xlsx")

# Calculate effect sizes: Standardized Mean Difference (between-subjects)
CRT_vs_ACC <- escalc(n1i = nEG, m1i = Mean_MMSE_CRT_Outcome, sd1i = SD_MMSE_CRT_Outcome, 
                     n2i = nCG, m2i = Mean_MMSE_Control_Outcome, sd2i = SD_MMSE_Control_Outcome, 
                     measure = "SMD", data = CRTvsACC)


# Run random-effects meta-analysis
rma_CRT_vs_ACC <- rma(yi, vi, data = CRT_vs_ACC, method = "REML")

# Display summary of results
summary(rma_CRT_vs_ACC)

# Random-Effects Model (k = 4; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -0.1671    0.3342    4.3342    2.5315   16.3342   
# 
# tau^2 (estimated amount of total heterogeneity): 0 (SE = 0.0623)
# tau (square root of estimated tau^2 value):      0
# I^2 (total heterogeneity / total variability):   0.00%
# H^2 (total variability / sampling variability):  1.00
# 
# Test for Heterogeneity:
#   Q(df = 3) = 1.4759, p-val = 0.6879
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub      
# 0.5964  0.1397  4.2706  <.0001  0.3227  0.8701  *** 

###moderate effect size CRT vs. ACC.

#### outlier check CRT vs ACC 
# Saving yi & vi in separate datasheet 
write.xlsx(CRT_vs_ACC, file = "CRT_vs_ACC_yi_vi.xlsx")
#outlier check
rstudent(rma_CRT_vs_ACC)
#     resid     se       z 
# 1 -0.2218 0.2953 -0.7510 
# 2  0.4026 0.3705  1.0866 
# 3 -0.2161 0.6004 -0.3600 
# 4  0.0167 0.2818  0.0591 

influence(rma_CRT_vs_ACC)
#    rstudent  dffits cook.d  cov.r tau2.del QE.del    hat  weight    dfbs inf 
# 1  -0.7510 -0.5361 0.2874 1.5096   0.0000 0.9119 0.3376 33.7587 -0.5361     
# 2   1.0866  0.4944 0.2444 1.2070   0.0000 0.2952 0.1715 17.1504  0.4944     
# 3  -0.3600 -0.0888 0.0079 1.0609   0.0000 1.3462 0.0574  5.7403 -0.0888     
# 4   0.0591  0.0517 0.0027 1.7652   0.0000 1.4724 0.4335 43.3507  0.0517

## no outlier

# Create forest plot
jpeg("Forest Plot - CRT vs. ACC.jpeg", width=600, height=400)
forest(rma_CRT_vs_ACC,
       slab = paste(CRTvsACC$Study.name, sep=""),
       psize = 1, lwd = 2,
       cex = .90,
       header = c("Reference", "Weight (%) SMDs [95% CI]"),
       addpred = TRUE,
       predstyle = "dist",
       showweights = TRUE)
usr <- par("usr")
x_pos <- usr[1] 
text(x_pos, -1.5, pos=4, cex=0.9, 
     bquote(paste("(Q = ", .(formatC(rma_CRT_vs_ACC$QE, digits=2, format="f")), 
                  ", df = ", .(rma_CRT_vs_ACC$k - rma_CRT_vs_ACC$p), 
                  ", p = ", .(formatC(rma_CRT_vs_ACC$QEp, digits=2, format="f")), 
                  "; ", I^2, " = ", .(formatC(rma_CRT_vs_ACC$I2, digits=1, format="f")), "%)")))
dev.off() 

#funnel plot
jpeg("funnel_plot_CRT_vs_ACC.jpeg")
funnel (rma_CRT_vs_ACC)
dev.off()

#Egger´s test
# k <10


#####################################################################
### 5) pairwise meta-analysis DCRT vs. ACC. (incl. sensitivity analysis) ###
#####################################################################

#data import
DCRTvsACC <- read_excel("DCRT_vs_ACC.xlsx")

# Calculate effect sizes: Standardized Mean Difference (between-subjects)
DCRT_vs_ACC <- escalc(n1i = nEG, m1i = Mean_MMSE_DCRT_Outcome, sd1i = SD_MMSE_DCRT_Outcome, 
                      n2i = nCG, m2i = Mean_MMSE_Control_Outcome, sd2i = SD_MMSE_Control_Outcome, 
                      measure = "SMD", data = DCRTvsACC)

# Run random-effects meta-analysis
rma_DCRT_vs_ACC <- rma(yi, vi, data = DCRT_vs_ACC, method = "REML")

# Display summary of results
summary(rma_DCRT_vs_ACC)

# Random-Effects Model (k = 9; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -5.2878   10.5755   14.5755   14.7344   16.9755   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0749 (SE = 0.0905)
# tau (square root of estimated tau^2 value):      0.2737
# I^2 (total heterogeneity / total variability):   43.19%
# H^2 (total variability / sampling variability):  1.76
# 
# Test for Heterogeneity:
#   Q(df = 8) = 14.6274, p-val = 0.0668
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.4429  0.1457  3.0404  0.0024  0.1574  0.7284  ** 

#### outlier check DCRT vs ACC 
# Saving yi & vi in separate datasheet 
write.xlsx(DCRT_vs_ACC, file = "DCRT_vs_ACC_yi_vi.xlsx")
#outlier check
rstudent(rma_DCRT_vs_ACC)
# resid     se       z 
# 1 -0.7062 0.5194 -1.3598 
# 2 -0.3013 0.4046 -0.7447 
# 3  0.0650 0.6566  0.0990 
# 4  0.4549 0.2603  1.7476 
# 5 -0.4949 0.4742 -1.0436 
# 6  0.5592 0.6099  0.9168 
# 7  0.9269 0.5699  1.6264 
# 8 -0.3639 0.4374 -0.8321 
# 9  0.2212 0.4729  0.4677 

influence(rma_DCRT_vs_ACC)
#    rstudent  dffits cook.d  cov.r tau2.del  QE.del    hat  weight    dfbs inf 
# 1  -1.3598 -0.3972 0.1515 1.0170   0.0647 11.9965 0.0821  8.2133 -0.4024     
# 2  -0.7447 -0.3327 0.1232 1.3063   0.0887 12.5807 0.1737 17.3675 -0.3370     
# 3   0.0990  0.0241 0.0006 1.1424   0.0876 14.6234 0.0538  5.3831  0.0236     
# 4   1.7476  1.1196 0.5893 0.8496   0.0189  9.8132 0.2006 20.0598  0.9722   * 
# 5  -1.0436 -0.3574 0.1273 1.1128   0.0743 12.7475 0.1051 10.5112 -0.3575     
# 6   0.9168  0.2375 0.0570 1.0901   0.0785 13.6737 0.0615  6.1460  0.2361     
# 7   1.6264  0.4120 0.1624 0.9750   0.0611 11.5834 0.0668  6.6758  0.4218     
# 8  -0.8321 -0.3286 0.1138 1.2168   0.0834 13.0162 0.1352 13.5155 -0.3288     
# 9   0.4677  0.1733 0.0337 1.2880   0.0962 14.3127 0.1213 12.1278  0.1726  

leave1out(rma_DCRT_vs_ACC)
#     estimate     se   zval   pval  ci.lb  ci.ub       Q     Qp   tau2      I2     H2 
# -1   0.4996 0.1469 3.4008 0.0007 0.2117 0.7875 11.9965 0.1007 0.0647 40.9297 1.6929 
# -2   0.4940 0.1665 2.9673 0.0030 0.1677 0.8204 12.5807 0.0830 0.0887 43.4763 1.7692 
# -3   0.4393 0.1557 2.8217 0.0048 0.1342 0.7445 14.6234 0.0411 0.0876 49.2492 1.9704 
# -4   0.3311 0.1343 2.4657 0.0137 0.0679 0.5942  9.8132 0.1994 0.0189 12.9727 1.1491 
# -5   0.4949 0.1537 3.2204 0.0013 0.1937 0.7961 12.7475 0.0785 0.0743 43.4388 1.7680 
# -6   0.4081 0.1521 2.6833 0.0073 0.1100 0.7062 13.6737 0.0573 0.0785 46.3143 1.8627 
# -7   0.3842 0.1438 2.6709 0.0076 0.1023 0.6661 11.5834 0.1151 0.0611 40.0337 1.6676 
# -8   0.4920 0.1607 3.0621 0.0022 0.1771 0.8070 13.0162 0.0717 0.0834 44.7971 1.8115 
# -9   0.4162 0.1653 2.5173 0.0118 0.0921 0.7402 14.3127 0.0459 0.0962 49.1149 1.9652 

# Create forest plot
jpeg("Forest Plot - DCRT vs. ACC.jpeg", width=600, height=500)
forest(rma_DCRT_vs_ACC,
       slab = paste(DCRTvsACC$Study.name, sep=""),
       psize = 1, lwd = 2,
       cex = .90,
       header = c("Reference", "Weight (%) SMDs [95% CI]"),
       addpred = TRUE,
       predstyle = "dist",
       showweights = TRUE)
usr <- par("usr")
x_pos <- usr[1] 
text(x_pos, -1.5, pos=4, cex=0.9, 
     bquote(paste("(Q = ", .(formatC(rma_DCRT_vs_ACC$QE, digits=2, format="f")), 
                  ", df = ", .(rma_DCRT_vs_ACC$k - rma_DCRT_vs_ACC$p), 
                  ", p = ", .(formatC(rma_DCRT_vs_ACC$QEp, digits=2, format="f")), 
                  "; ", I^2, " = ", .(formatC(rma_DCRT_vs_ACC$I2, digits=1, format="f")), "%)")))
dev.off() 

#funnel plot
jpeg("funnel_plot_DCRT_vs_ACC.jpeg")
funnel (rma_DCRT_vs_ACC)
dev.off()

#Egger´s test
#k <10

######################## sensitivity analysis #################################
# only studies comparing DCRT with TAU included

#data import
DCRTvsTAU <- read_excel("DCRT_vs_ACC_TAUonly.xlsx")

# Calculate effect sizes: Standardized Mean Difference (between-subjects)
DCRT_vs_TAU <- escalc(n1i = nEG, m1i = Mean_MMSE_DCRT_Outcome, sd1i = SD_MMSE_DCRT_Outcome, 
                      n2i = nCG, m2i = Mean_MMSE_Control_Outcome, sd2i = SD_MMSE_Control_Outcome, 
                      measure = "SMD", data = DCRTvsTAU)

# Run random-effects meta-analysis
rma_DCRT_vs_TAU <- rma(yi, vi, data = DCRT_vs_TAU, method = "REML")

# Display summary of results
summary(rma_DCRT_vs_TAU)

# Random-Effects Model (k = 7; tau^2 estimator: REML)
# 
# logLik  deviance       AIC       BIC      AICc   
# -4.2025    8.4050   12.4050   11.9886   16.4050   
# 
# tau^2 (estimated amount of total heterogeneity): 0.0893 (SE = 0.1078)
# tau (square root of estimated tau^2 value):      0.2989
# I^2 (total heterogeneity / total variability):   50.56%
# H^2 (total variability / sampling variability):  2.02
# 
# Test for Heterogeneity:
#   Q(df = 6) = 12.7475, p-val = 0.0472
# 
# Model Results:
#   
#estimate      se    zval    pval   ci.lb   ci.ub     
# 0.4955  0.1662  2.9811  0.0029  0.1697  0.8212  ** 

### outlier check DCRT vs TAU 
# Saving yi & vi in separate datasheet 
write.xlsx(DCRT_vs_TAU, file = "DCRT_vs_ACC_TAUonly_yi_vi.xlsx")
#outlier check
rstudent(rma_DCRT_vs_TAU)
#     resid     se       z 
# 1 -0.7763 0.5333 -1.4556 
# 2 -0.3793 0.4285 -0.8852 
# 3  0.3843 0.3682  1.0438 
# 4  0.5114 0.6318  0.8093 
# 5  0.8864 0.5839  1.5181 
# 6 -0.4387 0.4615 -0.9505 
# 7  0.1623 0.5188  0.3129 

influence(rma_DCRT_vs_TAU)
#    rstudent  dffits cook.d  cov.r tau2.del  QE.del    hat  weight    dfbs inf 
# 1  -1.4556 -0.4604 0.1995 1.0011   0.0733  9.7723 0.1012 10.1249 -0.4698     
# 2  -0.8852 -0.4440 0.2121 1.3253   0.0997  9.9682 0.2022 20.2168 -0.4471     
# 3   1.0438  0.6094 0.3180 1.1799   0.0721  8.8472 0.2297 22.9744  0.5957     
# 4   0.8093  0.2352 0.0566 1.1372   0.0974 11.9377 0.0768  7.6788  0.2324     
# 5   1.5181  0.4382 0.1816 0.9671   0.0713  9.9728 0.0831  8.3120  0.4514     
# 6  -0.9505 -0.4205 0.1842 1.2431   0.0966 10.6979 0.1611 16.1106 -0.4204     
# 7   0.3129  0.1119 0.0151 1.4457   0.1291 12.5627 0.1458 14.5825  0.1109  

## no outlier

# Create forest plot
jpeg("Forest Plot - DCRT vs. TAU.jpeg", width=600, height=500)
forest(rma_DCRT_vs_TAU,
       slab = paste(DCRTvsTAU$Study.name, sep=""),
       psize = 1, lwd = 2,
       cex = .90,
       header = c("Reference", "Weight (%) SMDs [95% CI]"),
       addpred = TRUE,
       predstyle = "dist",
       showweights = TRUE)
usr <- par("usr")
x_pos <- usr[1] 
text(x_pos, -1.5, pos=4, cex=0.9, 
     bquote(paste("(Q = ", .(formatC(rma_DCRT_vs_TAU$QE, digits=2, format="f")), 
                  ", df = ", .(rma_DCRT_vs_TAU$k - rma_DCRT_vs_TAU$p), 
                  ", p = ", .(formatC(rma_DCRT_vs_TAU$QEp, digits=2, format="f")), 
                  "; ", I^2, " = ", .(formatC(rma_DCRT_vs_TAU$I2, digits=1, format="f")), "%)")))
dev.off() 

#funnel plot
jpeg("funnel_plot_DCRT_vs_TAU.jpeg")
funnel (rma_DCRT_vs_TAU)
dev.off()

#Egger´s test
#k <10

#####################################################################
### 6) network meta-analysis                                      ###
#####################################################################

#data import
NMA <- read_excel("NMA.xlsx")

# Calculate effect sizes: Standardized Mean Difference (between-subjects)
NMA_ES <- escalc(n1i = nEG, m1i = Mean_EG, sd1i = SD_EG, 
                      n2i = nCG, m2i = Mean_CG, sd2i = SD_CG, 
                      measure = "SMD", data = NMA)

# Saving yi & vi in datasheet 
write.xlsx(NMA_ES, file = "NMA.xlsx")
NMA <- read_excel("NMA.xlsx")

#Create column for standard errors
NMA$SDi <- sqrt(NMA$vi)

#random effects model, ACC comparison
NMA_random_ACC <- netmeta(TE = yi,
                               seTE = SDi,
                               treat1 = Treatment1,
                               treat2 = Treatment2,
                               studlab = Study.name,
                               data = NMA,
                               sm = "SMD",
                               common = FALSE,
                               random = TRUE,
                               reference.group = "ACC",
                               sep.trts = " vs ",
                               tol.multiarm = 0.01, # tolerance threshold for the inconsistency of ES
                               tol.multiarm.se = 0.08)

summary(NMA_random_ACC)

# Original data (with adjusted standard errors for multi-arm studies):
#   
#                               treat1 treat2      TE   seTE seTE.adj narms multiarm
# Byeon (2018)                     CRT   DCRT -0.1821 0.4011   0.4457     2         
# Carcelen-Fraile et al. (2022)    ACC    CRT -0.4495 0.2404   0.3090     2         
# Diaz Baquero et al. (2022)       ACC   DCRT  0.2066 0.4283   0.4703     2         
# Han et al. (2017)                ACC   DCRT -0.1927 0.2174   0.2915     2         
# Kawashima et al. (2015)          ACC    CRT -0.9299 0.3372   0.3891     2         
# Lee et al. (2013)                ACC    CRT -0.3927 0.5829   0.7713     3        *
# Lee et al. (2013)                ACC   DCRT -0.5043 0.5651   0.7272     3        *
# Lee et al. (2013)                CRT   DCRT -0.1248 0.5569   0.7097     3        *
# Li et al. (2019)                 ACC   DCRT -0.7860 0.1757   0.2619     2         
# Lu et al. (2024)                 ACC   DCRT -0.0000 0.3563   0.4058     2         
# Oliveira et al. (2021)           ACC   DCRT -0.9673 0.5200   0.5550     2         
# Petri et al. (2024)              ACC   DCRT -1.3111 0.4929   0.5298     2         
# Prokopenko et al. (2019)         ACC   DCRT -0.1281 0.2865   0.3461     2         
# Savulich et al. (2017)           ACC   DCRT -0.6373 0.3163   0.3712     2         
# Tian et al. (2021)               ACC    CRT -0.6058 0.2121   0.2876     2         
# 
# Number of treatment arms (by study):
#                                 narms
# Byeon (2018)                      2
# Carcelen-Fraile et al. (2022)     2
# Diaz Baquero et al. (2022)        2
# Han et al. (2017)                 2
# Kawashima et al. (2015)           2
# Lee et al. (2013)                 3
# Li et al. (2019)                  2
# Lu et al. (2024)                  2
# Oliveira et al. (2021)            2
# Petri et al. (2024)               2
# Prokopenko et al. (2019)          2
# Savulich et al. (2017)            2
# Tian et al. (2021)                2
# 
# Results (random effects model):
#   
#                               treat1 treat2     SMD             95%-CI
# Byeon (2018)                     CRT   DCRT  0.0810 [-0.2925;  0.4545]
# Carcelen-Fraile et al. (2022)    ACC    CRT -0.5543 [-0.8759; -0.2328]
# Diaz Baquero et al. (2022)       ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Han et al. (2017)                ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Kawashima et al. (2015)          ACC    CRT -0.5543 [-0.8759; -0.2328]
# Lee et al. (2013)                ACC    CRT -0.5543 [-0.8759; -0.2328]
# Lee et al. (2013)                ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Lee et al. (2013)                CRT   DCRT  0.0810 [-0.2925;  0.4545]
# Li et al. (2019)                 ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Lu et al. (2024)                 ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Oliveira et al. (2021)           ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Petri et al. (2024)              ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Prokopenko et al. (2019)         ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Savulich et al. (2017)           ACC   DCRT -0.4733 [-0.7128; -0.2338]
# Tian et al. (2021)               ACC    CRT -0.5543 [-0.8759; -0.2328]
# 
# Number of studies: k = 13
# Number of pairwise comparisons: m = 15
# Number of treatments: n = 3
# Number of designs: d = 4
# 
# Random effects model
# 
# Treatment estimate (sm = 'SMD', comparison: other treatments vs 'ACC'):
#         SMD           95%-CI    z p-value
# ACC       .                .    .       .
# CRT  0.5543 [0.2328; 0.8759] 3.38  0.0007
# DCRT 0.4733 [0.2338; 0.7128] 3.87  0.0001
# 
# Quantifying heterogeneity / inconsistency:
#   tau^2 = 0.0377; tau = 0.1942; I^2 = 28% [0.0%; 62.7%]
# 
# Tests of heterogeneity (within designs) and inconsistency (between designs):
#                     Q d.f. p-value
# Total           16.66   12  0.1627
# Within designs  15.97    9  0.0675
# Between designs  0.68    3  0.8770
# 
# Details of network meta-analysis methods:
#   - Frequentist graph-theoretical approach
# - DerSimonian-Laird estimator for tau^2
# - Calculation of I^2 based on Q

####################### outlier check CRT vs ACC ###############################

#outlier check
(0.8759 - 0.2328) / 3.92 #0.1640561
0.5543  + (3.3* (0.1640561 * sqrt(4))) #upper:1.63707
0.5543  - (3.3* (0.1640561 * sqrt(4))) #lower:-0.5284703
#no outlier

####################### outlier check DCRT vs ACC ##############################

#outlier check
(0.7128 - 0.2338) / 3.92 #0.1221939
0.4733  + (3.3* (0.1221939 * sqrt(9))) #upper:1.68302
0.4733  - (3.3* (0.1221939 * sqrt(9))) #lower:-0.7364196
#no outlier

#random effects model, CRT comparison
NMA_random_CRT <- netmeta(TE = yi,
                          seTE = SDi,
                          treat1 = Treatment1,
                          treat2 = Treatment2,
                          studlab = Study.name,
                          data = NMA,
                          sm = "SMD",
                          common = FALSE,
                          random = TRUE,
                          reference.group = "CRT",
                          sep.trts = " vs ",
                          tol.multiarm = 0.01, # tolerance threshold for the inconsistency of ES
                          tol.multiarm.se = 0.08)

summary(NMA_random_CRT)

# Number of studies: k = 13
# Number of pairwise comparisons: m = 15
# Number of treatments: n = 3
# Number of designs: d = 4
# 
# Random effects model
# 
# Treatment estimate (sm = 'SMD', comparison: other treatments vs 'CRT'):
#           SMD             95%-CI     z p-value
# ACC  -0.5543 [-0.8759; -0.2328] -3.38  0.0007
# CRT        .                  .     .       .
# DCRT -0.0810 [-0.4545;  0.2925] -0.43  0.6707
# 
# Quantifying heterogeneity / inconsistency:
#   tau^2 = 0.0377; tau = 0.1942; I^2 = 28% [0.0%; 62.7%]
# 
# Tests of heterogeneity (within designs) and inconsistency (between designs):
#                     Q d.f. p-value
# Total           16.66   12  0.1627
# Within designs  15.97    9  0.0675
# Between designs  0.70    3  0.8741

####################### outlier check CRT vs DCRT ##############################

#outlier check
(0.2925 - (-0.4545)) / 3.92 #0.1905612
-0.0810  + (3.3* (0.1905612 * sqrt(2))) #upper:0.808331
-0.0810  - (3.3* (0.1905612 * sqrt(2))) #lower:-0.970331
#no outlier


#####################################################################
### 7) network graph & crosstabulation                            ###
#####################################################################
jpeg("network_graph.jpg", width = 800, height = 800, quality = 100)

netgraph(NMA_random_ACC,
         plastic = FALSE,                       
         thickness = "number.of.studies",       
         number.of.studies = TRUE,               
         points = TRUE,                          
         cex.points = 5,                         
         cex = 1.2,                              
         col = "darkblue",                      
         multiarm = TRUE)

dev.off()


#Number of pairwise comparisons per comparison design
pairs <- data.frame(NMA$Treatment1,NMA$Treatment2)
levels <- c("DCRT", "CRT","ACC")
pairs$Treatment1 <- factor(NMA$Treatment1, levels=levels)
pairs$Treatment2 <- factor(NMA$Treatment2, levels=levels)
tab <- table(pairs$Treatment1, pairs$Treatment2)
tab

#       DCRT CRT ACC
# DCRT    0   2   9
# CRT     0   0   4
# ACC     0   0   0


#####################################################################
### 8) forest plot                                                ###
#####################################################################

#ACC as comparison group
jpeg("Forest Plot - Network ACC.jpg", width = 400, height = 200, quality = 100)
labels2<-c("ACC", "CRT", "DCRT")
forest(NMA_random_ACC, 
       reference.group = "ACC",
       sortvar = -TE,
       xlim = c(-0.5, 1.5),
       smlab = paste(""),
       drop.reference.group = TRUE,
       label.left = "Favors ACC",
       label.right = "Favors (digital) remediation therapy",
       labels = labels2)
dev.off()


#####################################################################
### 9) Inconsistency checks                                       ###
#####################################################################
### 9.1) Net splitting
netsplit(NMA_random_ACC)

# comparison  k prop    nma  direct indir.    Diff     z p-value
# CRT vs ACC  4 0.86 0.5543  0.6039 0.2584  0.3455  0.74  0.4604
# DCRT vs ACC 9 0.93 0.4733  0.4459 0.8279 -0.3821 -0.81  0.4195
# CRT vs DCRT 2 0.29 0.0810 -0.1613 0.1787 -0.3400 -0.81  0.4195


### 9.2) Net Heat Plot

# Heatmap als JPEG speichern
jpeg("Network_Heatmap_ACC.jpeg", width = 800, height = 600, quality = 100)

netheat(NMA_random_ACC)

dev.off()


#####################################################################
### 10) Comparison adjusted funnel plot                           ###
#####################################################################
funnel(NMA_random_ACC,
       comparison = "all",          
       order = c("DCRT","CRT", "ACC"),
       adjust = TRUE)



