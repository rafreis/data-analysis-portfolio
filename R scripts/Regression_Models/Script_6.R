# Set the working directory
setwd("C:/Users/rafre/Dropbox/Fiverr/Trabalhos/2024/akendall066")

# Load the openxlsx library
library(openxlsx)
library(dplyr)

# Read data from a specific sheet of an Excel file
df <- read.xlsx("Copy of Examining the Impact of Attachment Style and the Strong Black Woman Schema on Marital Satisfaction_March 3_ 2025_08.04.xlsx", sheet = "Sheet0")

# Get rid of special characters
names(df) <- gsub(" ", "_", trimws(names(df)))
names(df) <- gsub("\\s+", "_", trimws(names(df), whitespace = "[\\h\\v\\s]+"))
names(df) <- gsub("\\(", "_", names(df))
names(df) <- gsub("\\)", "_", names(df))
names(df) <- gsub("\\-", "_", names(df))
names(df) <- gsub("/", "_", names(df))
names(df) <- gsub("\\\\", "_", names(df)) 
names(df) <- gsub("\\?", "", names(df))
names(df) <- gsub("\\'", "", names(df))
names(df) <- gsub("\\,", "_", names(df))
names(df) <- gsub("\\$", "", names(df))
names(df) <- gsub("\\+", "_", names(df))

# Trim all values

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
  }
  return(x)
}))

# Loop over each column in the dataframe
df <- data.frame(lapply(df, function(x) {
  if (is.character(x)) {
    # Apply trimws to each element in the column if it is character
    x <- trimws(x)
    # Convert character strings to lowercase
    x <- tolower(x)
  }
  return(x)
}))

clean_names <- function(names_vector) {
  names_vector <- gsub("_", "", names_vector)        # Remove underscores
  names_vector <- gsub("\\.{2,}", ".", names_vector) # Replace multiple dots with a single dot
  names_vector <- gsub("\\.+$", "", names_vector)    # Remove trailing dots
  return(names_vector)
}

# Apply to dataframe
colnames(df) <- clean_names(colnames(df))
colnames(df)

# Loop through each column to print the column name and unique values
for (col in names(df)) {
  cat("\nColumn:", col, "\n")
  print(unique(df[[col]]))
}

# Marital Satisfaction (DAS-7)
DAS7_items <- c(
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.1.Philosophy.of.life",
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.2.Aims.goals.and.things.believed.important",
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.3.Amount.of.time.spent.together",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.4.Have.a.stimulating.exchange.of.ideas",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.5.Calmly.discuss.something.together",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.6.Work.together.on.a.project"
)

attachment_anxiety <- c(
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.1.Im.afraid.that.I.will.lose.my.partners.love",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.3.I.often.worry.that.my.partner.will.not.want.to.stay.with.me",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.5.I.often.worry.that.my.partner.doesnt.really.love.me",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.7.I.worry.that.romantic.partners.won.t.care.about.me.as.much.as.I.care.about.them",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.9.I.often.wish.that.my.partners.feelings.for.me.were.as.strong.as.my.feelings.for.him.or.her",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.13.When.my.partner.is.out.of.sight.I.worry.that.he.or.she.might.become.interested.in",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.15.When.I.show.my.feelings.for.romantic.partners.Im.afraid.they.will.not.feel.the.same.about.me",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.29.Im.afraid.that.once.a.romantic.partner.gets.to.know.me.he.or.she.wont.like.who.I.really.am",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.33.I.worry.that.I.wont.measure.up.to.other.people"
)

avoidance <- c(
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.26.I.prefer.not.to.be.too.close.to.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.28.I.dont.feel.comfortable.opening.up.to.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.32.I.find.it.difficult.to.allow.myself.to.depend.on.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.35.I.prefer.not.to.show.a.partner.how.I.feel.deep.down"
)

sws <- c(
  "Please.indicate.how.much.you.experience.the.following.statements.1.I.try.to.present.an.image.of.strength",
  "Please.indicate.how.much.you.experience.the.following.statements.2.I.have.to.be.strong",
  "Please.indicate.how.much.you.experience.the.following.statements.3.I.feel.obligated.to.present.an.image.of.strength.at.work",
  "Please.indicate.how.much.you.experience.the.following.statements.4.I.feel.obligated.to.present.an.image.of.strength.for.my.family",
  "Please.indicate.how.much.you.experience.the.following.statements.6.I.keep.my.feelings.to.myself",
  "Please.indicate.how.much.you.experience.the.following.statements.7.My.tears.are.a.sign.of.weakness",
  "Please.indicate.how.much.you.experience.the.following.statements.8.I.keep.my.problems.bottled.up.inside",
  "Please.indicate.how.much.you.experience.the.following.statements.9.I.hide.my.stress",
  "Please.indicate.how.much.you.experience.the.following.statements.10.Expressing.emotions.is.difficult.for.me",
  "Please.indicate.how.much.you.experience.the.following.statements.11.It.s.hard.for.me.to.accept.help.from.others",
  "Please.indicate.how.much.you.experience.the.following.statements.13.I.wait.until.I.am.overwhelmed.to.ask.for.help",
  "Please.indicate.how.much.you.experience.the.following.statements.14.Asking.for.help.is.difficult.for.me",
  "Please.indicate.how.much.you.experience.the.following.statements.15.I.resist.help.to.prove.that.I.can.make.it.on.my.own",
  "Please.indicate.how.much.you.experience.the.following.statements.17.I.accomplish.my.goals.with.limited.resources",
  "Please.indicate.how.much.you.experience.the.following.statements.18.It.is.very.important.to.me.to.be.the.best.at.the.things.that.I.do",
  "Please.indicate.how.much.you.experience.the.following.statements.19.No.matter.how.hard.I.work.I.feel.like.I.should.do.more",
  "Please.indicate.how.much.you.experience.the.following.statements.20.I.put.pressure.on.myself.to.achieve.a.certain.level.of.accomplishment",
  "Please.indicate.how.much.you.experience.the.following.statements.21.I.take.on.roles.and.responsibilities.when.I.am.already.overwhelmed",
  "Please.indicate.how.much.you.experience.the.following.statements.24.I.feel.obligated.to.take.care.of.others",
  "Please.indicate.how.much.you.experience.the.following.statements.25.When.others.ask.for.my.help.I.say.yes.when.I.should.say.no",
  "Please.indicate.how.much.you.experience.the.following.statements.27.I.neglect.the.things.that.bring.me.joy",
  "Please.indicate.how.much.you.experience.the.following.statements.29.The.struggles.of.my.ancestors.require.me.to.be.strong",
  "Please.indicate.how.much.you.experience.the.following.statements.31.I.do.things.by.myself.without.asking.for.help",
  "Please.indicate.how.much.you.experience.the.following.statements.32.The.only.way.for.me.to.be.successful.is.to.work.hard",
  "Please.indicate.how.much.you.experience.the.following.statements.33.I.am.a.perfectionist",
  "Please.indicate.how.much.you.experience.the.following.statements.34.There.is.no.time.for.me.because.I.am.always.taking.care.of.others"
)

# ✅ Reverse-Scored Items (Extracted)
reverse_avoidance <- c(
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.4.Its.easy.for.me.to.be.affectionate.with.my.partner",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.6.I.find.it.easy.to.depend.on.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.8.I.feel.comfortable.depending.on.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.14.I.tell.my.partner.just.about.everything",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.22.I.find.it.relatively.easy.to.get.close.to.my.partner",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.30.I.am.very.comfortable.being.close.to.romantic.partners",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.34.I.feel.comfortable.sharing.my.private.thoughts.and.feelings.with.my.partner"
)

reverse_sws <- c(
  "Please.indicate.how.much.you.experience.the.following.statements.5.I.display.my.emotions.in.privacy",
  "Please.indicate.how.much.you.experience.the.following.statements.12.I.have.a.hard.time.trusting.others",
  "Please.indicate.how.much.you.experience.the.following.statements.16.If.I.want.things.done.right.I.do.them.myself",
  "Please.indicate.how.much.you.experience.the.following.statements.22.I.take.on.too.many.responsibilities.in.my.family",
  "Please.indicate.how.much.you.experience.the.following.statements.23.I.put.everyone.else.s.needs.before.mine",
  "Please.indicate.how.much.you.experience.the.following.statements.26.I.neglect.my.health.e.g.I.don.t.exercise.or.eat.like.I.should",
  "Please.indicate.how.much.you.experience.the.following.statements.28.I.feel.guilty.when.I.take.time.for.myself",
  "Please.indicate.how.much.you.experience.the.following.statements.30.I.keep.my.problems.to.myself.to.prevent.burdening.others",
  "Please.indicate.how.much.you.experience.the.following.statements.35.I.have.to.be.strong.because.I.am.a.woman"
)


agreement_scale <- c(
  "strongly disagree" = 1,
  "disagree" = 2,
  "slightly disagree" = 3,
  "neutral" = 4,
  "slightly agree" = 5,
  "agree" = 6,
  "strongly agree" = 7
)


reverse_agreement_scale <- c(
  "strongly agree" = 1,
  "agree" = 2,
  "slightly agree" = 3,
  "neutral" = 4,
  "slightly disagree" = 5,
  "disagree" = 6,
  "strongly disagree" = 7
)


das_scale <- c(
  "always disagree" = 1,
  "almost always disagree" = 2,
  "frequently disagree" = 3,
  "occasionally agree" = 4,
  "almost always agree" = 5,
  "always agree" = 6
)

das_frequency_scale <- c(
  "never" = 1,
  "less than once a month" = 2,
  "once or twice a month" = 3,
  "once or twice a week" = 4,
  "once a day" = 5,
  "more often" = 6
)


sws_scale <- c(
  "this is not true for me" = 1,
  "this is true for me rarely" = 2,
  "this is true for me sometimes" = 3,
  "this is true for me all the time" = 3,
  "this bothers me not at all" = 1,
  "this bothers me somewhat" = 2,
  "this bothers me very much" = 3
)


# Function to recode variables based on a given scale
recode_variables <- function(df, cols, scale) {
  df <- df %>%
    mutate(across(all_of(cols), ~ scale[.], .names = "{.col}.recoded"))
  return(df)
}

# Function to apply reverse scoring
reverse_score <- function(df, cols, scale) {
  df <- df %>%
    mutate(across(all_of(cols), ~ scale[.], .names = "{.col}.reversed"))
  return(df)
}


# ✅ Recoding DAS-7 Items (Marital Satisfaction)
df <- recode_variables(df, DAS7_items, das_scale)

# ✅ Recoding DAS Frequency Items
df <- recode_variables(df, 
                       c(
                         "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.4.Have.a.stimulating.exchange.of.ideas",
                         "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.5.Calmly.discuss.something.together",
                         "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.6.Work.together.on.a.project"
                       ), das_frequency_scale)

# ✅ Recoding Attachment Anxiety Items
df <- recode_variables(df, attachment_anxiety, agreement_scale)

# ✅ Recoding Attachment Avoidance (Non-Reverse Items)
avoidance_non_reverse <- setdiff(avoidance, reverse_avoidance)
df <- recode_variables(df, avoidance_non_reverse, agreement_scale)

# ✅ Recoding Superwoman Schema (Non-Reverse Items)
sws_non_reverse <- setdiff(sws, reverse_sws)
df <- recode_variables(df, sws_non_reverse, sws_scale)

# ✅ Reverse scoring (apply the reversed scale transformation)
df <- reverse_score(df, reverse_avoidance, reverse_agreement_scale)
df <- reverse_score(df, reverse_sws, sws_scale)


colnames(df)

DAS7_items_recoded <- c(
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.1.Philosophy.of.life.recoded",
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.2.Aims.goals.and.things.believed.important.recoded",
  "Most.persons.have.disagreements.in.their.relationships.Please.indicate.below.the.approximate.extent.of.agreement.or.disagreement.between.you.and.your.partner.for.each.item.on.the.following.list.3.Amount.of.time.spent.together.recoded",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.4.Have.a.stimulating.exchange.of.ideas.recoded",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.5.Calmly.discuss.something.together.recoded",
  "How.often.would.you.say.the.following.events.occur.between.you.and.your.mate.6.Work.together.on.a.project.recoded"
)

attachment_anxiety_recoded <- c(
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.1.Im.afraid.that.I.will.lose.my.partners.love.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.3.I.often.worry.that.my.partner.will.not.want.to.stay.with.me.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.5.I.often.worry.that.my.partner.doesnt.really.love.me.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.7.I.worry.that.romantic.partners.won.t.care.about.me.as.much.as.I.care.about.them.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.9.I.often.wish.that.my.partners.feelings.for.me.were.as.strong.as.my.feelings.for.him.or.her.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.13.When.my.partner.is.out.of.sight.I.worry.that.he.or.she.might.become.interested.in.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.15.When.I.show.my.feelings.for.romantic.partners.Im.afraid.they.will.not.feel.the.same.about.me.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.29.Im.afraid.that.once.a.romantic.partner.gets.to.know.me.he.or.she.wont.like.who.I.really.am.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.33.I.worry.that.I.wont.measure.up.to.other.people.recoded"
)

avoidance_recoded <- c(
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.26.I.prefer.not.to.be.too.close.to.romantic.partners.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.28.I.dont.feel.comfortable.opening.up.to.romantic.partners.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.32.I.find.it.difficult.to.allow.myself.to.depend.on.romantic.partners.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.35.I.prefer.not.to.show.a.partner.how.I.feel.deep.down.recoded",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.4.Its.easy.for.me.to.be.affectionate.with.my.partner.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.6.I.find.it.easy.to.depend.on.romantic.partners.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.8.I.feel.comfortable.depending.on.romantic.partners.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.14.I.tell.my.partner.just.about.everything.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.22.I.find.it.relatively.easy.to.get.close.to.my.partner.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.30.I.am.very.comfortable.being.close.to.romantic.partners.reversed",
  "Please.indicate.how.much.you.agree.or.disagree.with.the.statement.below.34.I.feel.comfortable.sharing.my.private.thoughts.and.feelings.with.my.partner.reversed"
)

sws <- c(
  "Please.indicate.how.much.you.experience.the.following.statements.1.I.try.to.present.an.image.of.strength.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.2.I.have.to.be.strong.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.3.I.feel.obligated.to.present.an.image.of.strength.at.work.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.4.I.feel.obligated.to.present.an.image.of.strength.for.my.family.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.6.I.keep.my.feelings.to.myself.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.7.My.tears.are.a.sign.of.weakness.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.8.I.keep.my.problems.bottled.up.inside.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.9.I.hide.my.stress.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.10.Expressing.emotions.is.difficult.for.me.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.11.It.s.hard.for.me.to.accept.help.from.others.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.13.I.wait.until.I.am.overwhelmed.to.ask.for.help.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.14.Asking.for.help.is.difficult.for.me.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.15.I.resist.help.to.prove.that.I.can.make.it.on.my.own.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.17.I.accomplish.my.goals.with.limited.resources.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.18.It.is.very.important.to.me.to.be.the.best.at.the.things.that.I.do.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.19.No.matter.how.hard.I.work.I.feel.like.I.should.do.more.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.20.I.put.pressure.on.myself.to.achieve.a.certain.level.of.accomplishment.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.21.I.take.on.roles.and.responsibilities.when.I.am.already.overwhelmed.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.24.I.feel.obligated.to.take.care.of.others.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.25.When.others.ask.for.my.help.I.say.yes.when.I.should.say.no.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.27.I.neglect.the.things.that.bring.me.joy.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.29.The.struggles.of.my.ancestors.require.me.to.be.strong.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.31.I.do.things.by.myself.without.asking.for.help.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.32.The.only.way.for.me.to.be.successful.is.to.work.hard.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.33.I.am.a.perfectionist.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.34.There.is.no.time.for.me.because.I.am.always.taking.care.of.others.recoded",
  "Please.indicate.how.much.you.experience.the.following.statements.5.I.display.my.emotions.in.privacy.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.12.I.have.a.hard.time.trusting.others.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.16.If.I.want.things.done.right.I.do.them.myself.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.22.I.take.on.too.many.responsibilities.in.my.family.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.23.I.put.everyone.else.s.needs.before.mine.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.26.I.neglect.my.health.e.g.I.don.t.exercise.or.eat.like.I.should.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.28.I.feel.guilty.when.I.take.time.for.myself.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.30.I.keep.my.problems.to.myself.to.prevent.burdening.others.reversed",
  "Please.indicate.how.much.you.experience.the.following.statements.35.I.have.to.be.strong.because.I.am.a.woman.reversed"
)

## RELIABILITY ANALYSIS

# Scale Construction

library(psych)
library(dplyr)

reliability_analysis <- function(data, scales) {
  # Initialize a dataframe to store the results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    SEM = numeric(),
    StDev = numeric(),
    ITC = numeric(),  # Added for item-total correlation
    Alpha = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Process each scale
  for (scale in names(scales)) {
    subset_data <- data[scales[[scale]]]
    alpha_results <- alpha(subset_data)
    alpha_val <- alpha_results$total$raw_alpha
    
    # Calculate statistics for each item in the scale
    for (item in scales[[scale]]) {
      item_data <- data[[item]]
      item_itc <- alpha_results$item.stats[item, "raw.r"] # Get ITC for the item
      item_mean <- mean(item_data, na.rm = TRUE)
      item_sem <- sd(item_data, na.rm = TRUE) / sqrt(sum(!is.na(item_data)))
      
      results <- rbind(results, data.frame(
        Variable = item,
        Mean = item_mean,
        SEM = item_sem,
        StDev = sd(item_data, na.rm = TRUE),
        ITC = item_itc,  # Include the ITC value
        Alpha = NA
      ))
    }
    
    # Calculate the mean score for the scale and add as a new column
    scale_mean <- rowMeans(subset_data, na.rm = TRUE)
    data[[scale]] <- scale_mean
    scale_mean_overall <- mean(scale_mean, na.rm = TRUE)
    scale_sem <- sd(scale_mean, na.rm = TRUE) / sqrt(sum(!is.na(scale_mean)))
    
    # Append scale statistics
    results <- rbind(results, data.frame(
      Variable = scale,
      Mean = scale_mean_overall,
      SEM = scale_sem,
      StDev = sd(scale_mean, na.rm = TRUE),
      ITC = NA,  # ITC is not applicable for the total scale
      Alpha = alpha_val
    ))
  }
  
  return(list(data_with_scales = data, statistics = results))
}

# Update scales list with corrected names
scales <- list(
  "SWS" = sws,
  "DAS7" = DAS7_items_recoded,
  "Attachment Anxiety" = attachment_anxiety_recoded,
  "Avoidance" = avoidance_recoded
)

alpha_results <- reliability_analysis(df, scales)

df <- alpha_results$data_with_scales
df_descriptives <- alpha_results$statistics

scales_freq <- c("X1.Are.you.currently.in.a.cisgendered.marriage"                 ,                                                                                                                                                                                     
                 "X2.Are.you.a.Black.woman"                                        ,                                                                                                                                                                                    
                  "X3.Was.your.gender.assigned.female.at.birth"                     ,                                                                                                                                                                                    
                 "X4.Are.you.between.the.ages.of.25.and.45"                          ,                                                                                                                                                                                  
                 "X5.Are.you.currently.or.have.you.previously.received.treatment.at.an.inpatient.or.residential.facility.due.to.concerns.with.drug.or.alcohol.use.e.g.a.facility.that.you.stay.at.overnight.and.are.continuously.monitored"   ,                         
                  "X6.Are.you.diagnosed.with.a.Substance.Use.Disorder.e.g.drug.or.alcohol.addiction"    ,                                                                                                                                                                
                  "What.is.your.age.in.years"                                                            ,                                                                                                                                                               
                  "Please.indicate.how.long.you.have.been.married"                                        ,                                                                                                                                                              
                  "Which.of.the.following.best.describes.your.gender.identity"                             ,                                                                                                                                                             
                  "What.is.your.race"                                                                       ,                                                                                                                                                            
                  "What.is.your.highest.level.of.education"                                                  ,                                                                                                                                                           
                  "What.is.your.relationship.status"                                                          ,                                                                                                                                                          
                  "Where.are.you.located"  )

create_frequency_tables <- function(data, categories) {
  all_freq_tables <- list() # Initialize an empty list to store all frequency tables
  
  # Iterate over each category variable
  for (category in categories) {
    # Ensure the category variable is a factor
    data[[category]] <- factor(data[[category]])
    
    # Calculate counts
    counts <- table(data[[category]])
    
    # Create a dataframe for this category
    freq_table <- data.frame(
      "Category" = rep(category, length(counts)),
      "Level" = names(counts),
      "Count" = as.integer(counts),
      stringsAsFactors = FALSE
    )
    
    # Calculate and add percentages
    freq_table$Percentage <- (freq_table$Count / sum(freq_table$Count)) * 100
    
    # Add the result to the list
    all_freq_tables[[category]] <- freq_table
  }
  
  # Combine all frequency tables into a single dataframe
  do.call(rbind, all_freq_tables)
}

df_freq <- create_frequency_tables(df, scales_freq)

library(dplyr)
library(stringr)

# Function to convert marriage duration to numeric years
convert_to_years <- function(x) {
  if (is.na(x) || x == "") {
    return(NA)  # Handle missing values safely
  }
  
  x <- tolower(x) # Convert to lowercase for consistency
  
  # Direct numeric conversion for simple cases
  if (str_detect(x, "^\\d+(\\.\\d+)?$")) {
    return(as.numeric(x))
  }
  
  # Convert spelled-out numbers (basic)
  x <- str_replace_all(x, c("one" = "1", "two" = "2", "three" = "3", "four" = "4",
                            "five" = "5", "six" = "6", "seven" = "7", "eight" = "8",
                            "nine" = "9", "ten" = "10", "eleven" = "11", "twelve" = "12"))
  
  # Extract numerical values
  years <- as.numeric(str_extract(x, "\\b\\d+(\\.\\d+)?(?=\\s?(yr|yrs|year|years|$))"))
  months <- as.numeric(str_extract(x, "\\b\\d+(?=\\s?(month|months))")) # Extract months if mentioned
  
  # Convert months to years
  if (!is.na(months)) {
    if (is.na(years)) { years <- 0 } # If no years are mentioned, assume 0
    years <- years + (months / 12)
  }
  
  return(years)
}

# Apply function to the marriage duration column safely
df <- df %>%
  mutate(Marriage_Years = sapply(`Please.indicate.how.long.you.have.been.married`, convert_to_years))

# Check results
summary(df$Marriage_Years)

library(moments)

# Function to calculate descriptive statistics
calculate_descriptive_stats <- function(data, desc_vars) {
  # Initialize an empty dataframe to store results
  results <- data.frame(
    Variable = character(),
    Mean = numeric(),
    Median = numeric(),
    SEM = numeric(),  # Standard Error of the Mean
    SD = numeric(),   # Standard Deviation
    Skewness = numeric(),
    Kurtosis = numeric(),
    stringsAsFactors = FALSE
  )
  
  # Iterate over each variable and calculate statistics
  for (var in desc_vars) {
    variable_data <- data[[var]]
    mean_val <- mean(variable_data, na.rm = TRUE)
    median_val <- median(variable_data, na.rm = TRUE)
    sem_val <- sd(variable_data, na.rm = TRUE) / sqrt(length(na.omit(variable_data)))
    sd_val <- sd(variable_data, na.rm = TRUE)
    skewness_val <- skewness(variable_data, na.rm = TRUE)
    kurtosis_val <- kurtosis(variable_data, na.rm = TRUE)
    
    # Append the results for each variable to the results dataframe
    results <- rbind(results, data.frame(
      Variable = var,
      Mean = mean_val,
      Median = median_val,
      SEM = sem_val,
      SD = sd_val,
      Skewness = skewness_val,
      Kurtosis = kurtosis_val
    ))
  }
  
  return(results)
}

vars <- c("What.is.your.age.in.years", "Marriage_Years", "SWS" ,                                                                                                                                                                                                                                                
          "DAS7"                                                ,                                                                                                                                                                                                
          "Attachment Anxiety"                                           ,                                                                                                                                                                                               
          "Avoidance"  )

df$What.is.your.age.in.years <- as.numeric(as.character(df$What.is.your.age.in.years))
df$Marriage_Years <- as.numeric(as.character(df$Marriage_Years))
colnames(df)

df_descriptive_stats <- calculate_descriptive_stats(df, vars)

library(ggplot2)
library(tidyr)
library(stringr)

create_boxplots <- function(df, vars) {
  # Ensure that vars are in the dataframe
  df <- df[, vars, drop = FALSE]
  
  # Reshape the data to a long format
  long_df <- df %>%
    gather(key = "Variable", value = "Value")
  
  # Create side-by-side boxplots for each variable
  p <- ggplot(long_df, aes(x = Variable, y = Value)) +
    geom_boxplot() +
    stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "red") +
    labs(title = "Boxplots", x = "Variable", y = "Value") +
    theme(axis.text.x = element_text(angle = 0, hjust = 1, vjust = 1)) +
    scale_x_discrete(labels = function(x) str_wrap(x, width = 10))
  
  # Print the plot
  print(p)
  
  # Save the plot
  ggsave(filename = "boxplots3.png", plot = p, width = 10, height = 6)
}

vars <- c("SWS" ,                                                                                                                                                                                                                                                
          "DAS7"                                                ,                                                                                                                                                                                                
          "Attachment Anxiety"                                           ,                                                                                                                                                                                               
          "Avoidance"  )

create_boxplots(df, vars)

vars <- c("What.is.your.age.in.years" )
create_boxplots(df, vars)

vars <- c( "Marriage_Years")
create_boxplots(df, vars)

library(dplyr)
library(dplyr)
library(stringr)

calculate_adjusted_z_scores <- function(data, vars, id_var, z_threshold = 3.5) {
  # Prepare data by handling NA values and calculating Adjusted Z-scores
  z_score_data <- data %>%
    mutate(across(all_of(vars), ~ as.numeric(.))) %>%
    mutate(across(all_of(vars), ~ replace(., is.na(.), median(., na.rm = TRUE)))) %>% # Replace NA with median
    mutate(across(all_of(vars), ~ (0.6745 * (. - median(., na.rm = TRUE))) / ifelse(mad(., na.rm = TRUE) == 0, 1, mad(., na.rm = TRUE)), .names = "z_{.col}")) # Compute Adjusted Z-score

# Include original raw scores and Adjusted Z-scores
z_score_data <- z_score_data %>%
  select(!!sym(id_var), all_of(vars), starts_with("z_"))

# Add a column to flag outliers based on the specified Adjusted Z-score threshold
z_score_data <- z_score_data %>%
  mutate(across(starts_with("z_"), ~ ifelse(abs(.) > z_threshold, "Outlier", "Normal"), .names = "flag_{.col}"))

return(z_score_data)
}

vars <- c("What.is.your.age.in.years", "Marriage_Years", "SWS" ,                                                                                                                                                                                                                                                
          "DAS7"                                                ,                                                                                                                                                                                                
          "Attachment Anxiety"                                           ,                                                                                                                                                                                               
          "Avoidance"  )

# Run the function on your dataframe
df_zscores <- calculate_adjusted_z_scores(df, vars, "Response.ID")

library(dplyr)
library(tidyr)
library(stringr)

# Function to remove outlier values based on Z-score table
remove_outliers_based_on_z_scores <- function(data, z_score_data, id_var) {
  # Reshape the z-score data to long format
  z_score_data_long <- z_score_data %>%
    pivot_longer(cols = starts_with("z_"), names_to = "Variable", values_to = "z_value") %>%
    pivot_longer(cols = starts_with("flag_z_"), names_to = "Flag_Variable", values_to = "Flag_Value") %>%
    filter(str_replace(Flag_Variable, "flag_", "") == Variable) %>%
    select(!!sym(id_var), Variable, Flag_Value)
  
  # Remove flagged outliers from the original data
  for (i in 1:nrow(z_score_data_long)) {
    if (z_score_data_long$Flag_Value[i] == "Outlier") {
      col_name <- gsub("z_", "", z_score_data_long$Variable[i]) # Get original column name
      data <- data %>%
        mutate(!!sym(col_name) := ifelse(!!sym(id_var) == z_score_data_long[[id_var]][i], NA, !!sym(col_name)))
    }
  }
  
  return(data)
}

# Example usage
df_nooutliers <- remove_outliers_based_on_z_scores(df, df_zscores, id_var = "Response.ID")

df_descriptive_stats2 <- calculate_descriptive_stats(df_nooutliers, vars)

## CORRELATION ANALYSIS

calculate_correlation_matrix <- function(data, variables, method = "pearson") {
  # Ensure the method is valid
  if (!method %in% c("pearson", "spearman")) {
    stop("Method must be either 'pearson' or 'spearman'")
  }
  
  # Subset the data for the specified variables
  data_subset <- data[variables]
  
  # Calculate the correlation matrix
  corr_matrix <- cor(data_subset, method = method, use = "complete.obs")
  
  # Initialize a matrix to hold formatted correlation values
  corr_matrix_formatted <- matrix(nrow = nrow(corr_matrix), ncol = ncol(corr_matrix))
  rownames(corr_matrix_formatted) <- colnames(corr_matrix)
  colnames(corr_matrix_formatted) <- colnames(corr_matrix)
  
  # Calculate p-values and apply formatting with stars
  for (i in 1:nrow(corr_matrix)) {
    for (j in 1:ncol(corr_matrix)) {
      if (i == j) {  # Diagonal elements (correlation with itself)
        corr_matrix_formatted[i, j] <- format(round(corr_matrix[i, j], 3), nsmall = 3)
      } else {
        p_value <- cor.test(data_subset[[i]], data_subset[[j]], method = method)$p.value
        stars <- ifelse(p_value < 0.01, "***", 
                        ifelse(p_value < 0.05, "**", 
                               ifelse(p_value < 0.1, "*", "")))
        corr_matrix_formatted[i, j] <- paste0(format(round(corr_matrix[i, j], 3), nsmall = 3), stars)
      }
    }
  }
  
  return(corr_matrix_formatted)
}

correlation_matrix <- calculate_correlation_matrix(df, vars, method = "pearson")

# OLS Regression
library(broom)

fit_ols_and_format <- function(data, predictors, response_vars, save_plots = FALSE) {
  ols_results_list <- list()
  
  for (response_var in response_vars) {
    # Construct the formula dynamically
    formula <- as.formula(paste(response_var, "~", paste(predictors, collapse = " + ")))
    
    # Fit the OLS multiple regression model
    lm_model <- lm(formula, data = data)
    
    # Get the summary of the model
    model_summary <- summary(lm_model)
    
    # Extract the R-squared, F-statistic, and p-value
    r_squared <- model_summary$r.squared
    f_statistic <- model_summary$fstatistic[1]
    p_value <- pf(f_statistic, model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE)
    
    # Print the summary of the model for fit statistics
    print(summary(lm_model))
    
    # Extract the tidy output and assign it to ols_results
    ols_results <- broom::tidy(lm_model) %>%
      mutate(ResponseVariable = response_var, R_Squared = r_squared, F_Statistic = f_statistic, P_Value = p_value)
    
    # Optionally print the tidy output
    print(ols_results)
    
    # Generate and save residual plots
    if (save_plots) {
      plot_name_prefix <- gsub(" ", "_", response_var)
      
      jpeg(paste0(plot_name_prefix, "_Residuals_vs_Fitted.jpeg"))
      plot(residuals(lm_model) ~ fitted(lm_model), main = "Residuals vs Fitted", xlab = "Fitted values", ylab = "Residuals")
      dev.off()
      
      jpeg(paste0(plot_name_prefix, "_QQ_Plot.jpeg"))
      qqnorm(residuals(lm_model), main = "Q-Q Plot")
      qqline(residuals(lm_model))
      dev.off()
    }
    
    # Store the results in a list
    ols_results_list[[response_var]] <- ols_results
  }
  
  return(ols_results_list)
}

colnames(df_nooutliers)

dvs_list <- "DAS7"

ivs <- c("SWS"           ,                                                                                                                                                                                                                                      
                                                                                                                                                                                                                  
         "Attachment"      ,                                                                                                                                                                                                                                    
         "Avoidance" )

df_model_results <- fit_ols_and_format(
  data = df_nooutliers,
  predictors = ivs,  # Replace with actual predictor names
  response_vars = dvs_list,       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe
df_model_results <- bind_rows(df_model_results)

ivs_control <- c("SWS"           ,                                                                                                                                                                                                                                      
                                                                                                                                                                                                                                             
         "Attachment"      ,                                                                                                                                                                                                                                    
         "Avoidance",
         "Marriage_Years" ,
         "What.is.your.age.in.years"
         
         )

df_model_results_controlled <- fit_ols_and_format(
  data = df_nooutliers,
  predictors = ivs_control,  # Replace with actual predictor names
  response_vars = dvs_list,       # Replace with actual response variable names
  save_plots = TRUE
)

# Combine the results into a single dataframe
df_model_results_controlled <- bind_rows(df_model_results_controlled)


## Export Results

library(openxlsx)

save_apa_formatted_excel <- function(data_list, filename) {
  wb <- createWorkbook()  # Create a new workbook
  
  for (i in seq_along(data_list)) {
    # Define the sheet name
    sheet_name <- names(data_list)[i]
    if (is.null(sheet_name)) sheet_name <- paste("Sheet", i)
    addWorksheet(wb, sheet_name)  # Add a sheet to the workbook
    
    # Convert matrices to data frames, if necessary
    data_to_write <- if (is.matrix(data_list[[i]])) as.data.frame(data_list[[i]]) else data_list[[i]]
    
    # Include row names as a separate column, if they exist
    if (!is.null(row.names(data_to_write))) {
      data_to_write <- cbind("Index" = row.names(data_to_write), data_to_write)
    }
    
    # Write the data to the sheet
    writeData(wb, sheet_name, data_to_write)
    
    # Define styles
    header_style <- createStyle(textDecoration = "bold", border = "TopBottom", borderColour = "black", borderStyle = "thin")
    bottom_border_style <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin")
    
    # Apply styles
    addStyle(wb, sheet_name, style = header_style, rows = 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    addStyle(wb, sheet_name, style = bottom_border_style, rows = nrow(data_to_write) + 1, cols = 1:ncol(data_to_write), gridExpand = TRUE)
    
    # Set column widths to auto
    setColWidths(wb, sheet_name, cols = 1:ncol(data_to_write), widths = "auto")
  }
  
  # Save the workbook
  saveWorkbook(wb, filename, overwrite = TRUE)
}

# Example usage
data_list <- list(
  "Final Data" = df_zscores,
  "Frequency Table" = df_freq, 
  "Descriptive Stats" = df_descriptive_stats2, 
  "Reliability" = df_descriptives,
  "Correlation" = correlation_matrix,
  "OLS Regression" = df_model_results,
  "OLS Regression Contr" = df_model_results_controlled
)

# Save to Excel with APA formatting
save_apa_formatted_excel(data_list, "APA_Formatted_Tables2.xlsx")