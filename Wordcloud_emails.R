####################
## TEXT ANALYTICS ##
####################
# 1. Load an excel file with ~200 emails
# 2. Clean the corpus of the emails
# 3. Extract word frequencies
# 4. Create a word cloud with the frequencies 

rm(list=ls(all=TRUE))  # Clean up the memory of R session
setwd("~/Desktop/leaderconn")  #set working directory  #MAC

#install.packages(c("tm", "wordcloud", "SnowballC"))

library(tm)
library(SnowballC)
library(wordcloud)

library(xlsx)
leaders <- read.xlsx("emails.xlsx", sheetName="Raw Ideas List", header=T); names(leaders)
rawcomments <- leaders$Raw.Comment
head(rawcomments)
strwrap(rawcomments[1])

comments = Corpus(VectorSource(rawcomments)) ; comments

badWordsList  <- read.csv("removewords.txt",header = FALSE,stringsAsFactors = FALSE)
badWords <- badWordsList[,1]

cleanCorpus  <- function(corpus){
  require(tm)
  corpustmp = tm_map(corpus, tolower) #convert to lower case
  corpustmp = tm_map(corpustmp, PlainTextDocument)
  corpustmp = tm_map(corpustmp, removePunctuation) #remove punctuation
  corpustmp = tm_map(corpustmp,stripWhitespace) # clean up any trailing whitespace
  corpustmp = tm_map(corpustmp, removeNumbers) 
  corpustmp = tm_map(corpustmp, removeWords, stopwords("english")) #Remove stop words
  # dictCorpus <- corpustmp
  corpustmp <- tm_map(corpustmp, stemDocument) #stem document 
  # corpustmp <- tm_map(corpustmp, stemCompletion, dictionary=dictCorpus) #stem completion  //It's causing problems  corpustmp = tm_map(corpustmp, removeWords, badWords) 
  corpustmp = tm_map(corpustmp, removeWords, badWords) 
  
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "employe", replacement = "employees")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = c("custom","customer","customers","passeng"), replacement = "customers")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "issu", replacement = "issues")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "wifi", replacement = "WIFI")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "servic", replacement = "services")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "departur", replacement = "departures")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "bag", replacement = "bags")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "chang", replacement = "changes")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "compani", replacement = "companies")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = c("tech","ops"), replacement = "tech-ops")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "outsourc", replacement = "outsourcing")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "equip", replacement = "equipment")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "manag", replacement = "management")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "improv", replacement = "improvement")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "oper", replacement = "operations")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "requir", replacement = "requirements")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "communic", replacement = "communication")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "delay", replacement = "delays")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "charg", replacement = "charges")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "upgrad", replacement = "upgrades")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "avail", replacement = "availability")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "seat", replacement = "seating")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "confus", replacement = "confussion")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "secur", replacement = "security")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "cost", replacement = "costs")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "mainten", replacement = "maintenance")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "schedul", replacement = "scheduling")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "staf", replacement = "staffing")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "gse", replacement = "GSE")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "inform", replacement = "information")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "polici", replacement = "policies")
  corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "eta", replacement = "ETA")
  
  #corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "ue", replacement = "ues")
  #corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "+er", replacement = "ers")
  #corpustmp <- tm_map(corpustmp, content_transformer(gsub),pattern = "ee", replacement = "ees")
  return(corpustmp)
}

comments_clean <- cleanCorpus(comments)

#Extract the word frequencies
myDTM = TermDocumentMatrix(comments_clean) #,control=list(wordLenghts=c(4,25)))
myDTM   # we have 3131 terms that appear at least once.

#Remove sparse terms
#myDTM1 <- removeSparseTerms(myDTM, 0.99); myDTM1  # 139 terms
#myDTM2 <- removeSparseTerms(myDTM, 0.98); myDTM2  # 41 terms
#myDTM3 <- removeSparseTerms(myDTM, 0.97); myDTM3  # 17 terms

m = as.matrix(myDTM)

#Word Cloud
v = sort(rowSums(m), decreasing = TRUE)
set.seed(1231)
wordColors <- brewer.pal(8,"Dark2")
wordcloud(names(v), v , min.freq=15,colors=wordColors)

#Word frequencies
freq <- rowSums(m)
ord <- order(freq, decreasing=TRUE)
freq[head(ord)]
freq[tail(ord)]

d <- data.frame(word = names(v),freq=v)
table(d$word)

freq.df = data.frame(freq)
dim(freq.df)
write.csv(freq.df, file="WORDS_freq.csv")


