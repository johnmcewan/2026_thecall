#Topic modeling w/o named entity compounding before tokenizing

# install ldatuning from GitHub (CRAN version incompatible with R 4.5.x)
install.packages("remotes")
remotes::install_github("nikita-moor/ldatuning")

# data source from MNF project at IDRH
library(topicmodels)
library(ldatuning)
library(quanteda)
library(tidyverse)

# import data
# set your working directory before importing
reviews_raw <- read.csv("filelist_plaintext.csv")

# take a look at the data
View(reviews_raw)

# drop the one NA row (your data has 'id' and 'plain_text' columns, not 'Rating'/'Review')
reviews_clean <- reviews_raw %>% filter(!is.na(plain_text))

# create a corpus to preprocess the data
reviews_corpus <- corpus(reviews_clean, text_field = "plain_text")

# take a look at the corpus
summary(reviews_corpus)

# preprocess
reviews_processed <- tokens(reviews_corpus,
                            remove_punct = TRUE,
                            remove_numbers = TRUE,
                            remove_separators = TRUE,
                            remove_url = TRUE,
                            remove_symbols = TRUE)

# lowercase
reviews_processed <- tokens_tolower(reviews_processed)

# stopword removal
reviews_processed <- tokens_remove(reviews_processed, stopwords("en", "smart"))

# remove domain-specific noise words
# after initial LDA, add any noisy terms you spot here
reviews_processed <- tokens_remove(reviews_processed,
                                   c("chiefs", "monday", "nite", "footballers",
                                     "just", "like", "will", "one", "can", "now"))

# create a dfm/dtm for topic modeling
reviews_dfm <- dfm(reviews_processed) %>%
  dfm_trim(min_termfreq = 2)  # drop words appearing only once

# convert to format topicmodels expects
reviews_dtm <- convert(reviews_dfm, to = "topicmodels")

# topic modeling experiment with k = 7
reviews_lda <- LDA(reviews_dtm, k = 7, method = "Gibbs",
                   control = list(seed = 42, burnin = 500, iter = 1000))

# check the top 10 terms associated with each topic
terms(reviews_lda, 10)

# find the best number of k
proper_k <- FindTopicsNumber(reviews_dtm,
                             topics = seq(from = 3,   # start at k = 3
                                          to = 20,    # end by k = 20
                                          by = 3),    # jump up by 3 each time
                             metrics = c("Griffiths2004",
                                         "CaoJuan2009",
                                         "Arun2010",
                                         "Deveaud2014"),
                             control = list(seed = 42)
)

# plot the best k
FindTopicsNumber_plot(proper_k)

# Griffiths & Steyvers (2004) measures perplexity — lower is better fit
# Cao & Juan (2009) measures topic coherence via word co-occurrence
# Arun et al. (2010) measures coherence by word co-occurrence in documents
# Deveaud et al. (2014) measures coherence against a reference corpus
# Maximize: Griffiths2004, Deveaud2014 | Minimize: CaoJuan2009, Arun2010

# refit using best k suggested by the plot (adjust k as needed)
set.seed(42)
reviews_lda_4 <- LDA(reviews_dtm, k = 4, method = "Gibbs",
                     control = list(seed = 42, burnin = 500, iter = 1000))

# print the modeling results
lda_4_terms <- terms(reviews_lda_4, 10)
View(lda_4_terms)

# view the distribution/proportion of each topic across documents
posterior(reviews_lda_4)$topics

# view keyword probabilities for each topic
lda_4_terms_prob <- posterior(reviews_lda_4)$terms

# topic 1
head(sort(lda_4_terms_prob[1,], decreasing = TRUE), 10)

# topic 2
head(sort(lda_4_terms_prob[2,], decreasing = TRUE), 10)

# topic 3
head(sort(lda_4_terms_prob[3,], decreasing = TRUE), 10)

# topic 4
head(sort(lda_4_terms_prob[4,], decreasing = TRUE), 10)

####################### (optional)
# create a dataframe of all top keywords and their probabilities per topic
lda_4_terms_df <- as.data.frame(lda_4_terms_prob) %>%
  rownames_to_column("topic_id") %>%
  pivot_longer(-topic_id, names_to = "word", values_to = "prob")

# select top 10 terms per topic
lda_4_top_terms <- lda_4_terms_df %>%
  group_by(topic_id) %>%
  slice_max(order_by = prob, n = 10) %>%
  arrange(topic_id, desc(prob))

# take a look at the result
View(lda_4_top_terms)