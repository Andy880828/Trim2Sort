# library list
library(dada2); packageVersion("dada2")
library(ShortRead)
library(Biostrings)
library(argparser)

parser <- arg_parser("process input")
parser <- add_argument(parser, "--inpath", help="input folder route")
parser <- add_argument(parser, "--cutadapt", help="cutadapt route")
parser <- add_argument(parser, "--FWD", help="forword primer", default="GTCGGTAAAACTCGTGCCAGC")
parser <- add_argument(parser, "--REV", help="reverse primer", default="CATAGTGGGGTATCTAATCCCAGTTTG")
args <- parse_args(parser, commandArgs(trailingOnly = TRUE))

# routes
#cat("Please choose your folder：")
#inpath <- "D:/test_sample/" #測試用
#inpath <- choose.dir() #實際用
inpath <- args$inpath
list.files(inpath)

#cutadapt <- "D:/Trim2Sort_ver3.0.0/cutadapt.exe"
cutadapt <- args$cutadapt
system2(cutadapt, args = "--version")

# Forward and reverse fastq filenames have format: SAMPLENAME.R1.clean.fastq and SAMPLENAME.R2.clean.fastq
fnFs <- sort(list.files(inpath, pattern=".R1.clean.fastq", full.names = TRUE))
fnRs <- sort(list.files(inpath, pattern=".R2.clean.fastq", full.names = TRUE))

# Extract sample names, assuming filenames have format: SAMPLENAME.XXX.fastq
sample.names <- sapply(strsplit(basename(fnFs), ".R1.clean.fastq"), `[`, 1)
sample.cnt <- length(sample.names)
#print(sample.names)
#print(sample.cnt)

# # visualizing the quality profiles of the forward reads
# plotQualityProfile(fnFs[1:2])
# #In gray-scale is a heat map of the frequency of each quality score at each base position.
# #The mean quality score at each position is shown by the green line.
# #The quartiles of the quality score distribution by the orange lines.
# #The red line shows the scaled proportion of reads that extend to at least that position (this is more useful for other sequencing technologies, as Illumina reads are typically all the same length, hence the flat red line).
# 
# # visualizing the quality profiles of the reverse reads
# plotQualityProfile(fnRs[1:2])

#FWD <- "GTCGGTAAAACTCGTGCCAGC"
FWD <- args$FWD
#REV <- "CATAGTGGGGTATCTAATCCCAGTTTG"
REV <- args$REV

path.cut <- file.path(inpath, "cutadapt")
if(!dir.exists(path.cut)) dir.create(path.cut)
fnFs.cut <- file.path(path.cut, basename(fnFs))
fnRs.cut <- file.path(path.cut, basename(fnRs))

FWD.RC <- dada2:::rc(FWD)
REV.RC <- dada2:::rc(REV)
# Trim FWD and the reverse-complement of REV off of R1 (forward reads)
R1.flags <- paste("-g", FWD, "-a", REV.RC) 
# Trim REV and the reverse-complement of FWD off of R2 (reverse reads)
R2.flags <- paste("-G", REV, "-A", FWD.RC) 
# Run Cutadapt
for(i in seq_along(fnFs)) {
  system2(cutadapt, args = c(R1.flags, R2.flags, "-n", 2, # -n 2 required to remove FWD and REV from reads
                             "-o", fnFs.cut[i], "-p", fnRs.cut[i], # output files
                             fnFs[i], fnRs[i])) # input files
}


fnFcs <- sort(list.files(path.cut, pattern=".R1.clean.fastq", full.names = TRUE))
fnRcs <- sort(list.files(path.cut, pattern=".R2.clean.fastq", full.names = TRUE))


# Place filtered files in filtered/ subdirectory
filtFs <- file.path(inpath, "filtered", paste0(sample.names, "_F_filt.fastq.gz"))
filtRs <- file.path(inpath, "filtered", paste0(sample.names, "_R_filt.fastq.gz"))
names(filtFs) <- sample.names
names(filtRs) <- sample.names


out <- filterAndTrim(fnFcs, filtFs, fnRcs, filtRs, truncLen=c(160,100),
                     maxN=0, maxEE=c(1,1), truncQ=2, rm.phix=TRUE,
                     compress=TRUE, multithread=FALSE) # On Windows set multithread=FALSE
head(out)

# total bases for learning the error rates
errF <- learnErrors(filtFs, multithread=TRUE)
errR <- learnErrors(filtRs, multithread=TRUE)
plotErrors(errF, nominalQ=TRUE)

# apply the core sample inference algorithm to the filtered and trimmed sequence data
dadaFs <- dada(filtFs, err=errF, multithread=TRUE)
dadaRs <- dada(filtRs, err=errR, multithread=TRUE)
dadaFs[[1]]

mergers <- mergePairs(dadaFs, filtFs, dadaRs, filtRs, verbose=TRUE)
# Inspect the merger data.frame from the first sample
head(mergers[[1]])

# Construct sequence table
seqtab <- makeSequenceTable(mergers)
dim(seqtab)
# Inspect distribution of sequence lengths
table(nchar(getSequences(seqtab)))


# Remove chimeras
seqtab.nochim <- removeBimeraDenovo(seqtab, method="consensus", multithread=TRUE, verbose=TRUE)
dim(seqtab.nochim)

# Detect chimeras rate
sum(seqtab.nochim)/sum(seqtab)


# See the reduction of every step
# If processing a single sample, remove the sapply calls: e.g. replace sapply(dadaFs, getN) with getN(dadaFs)
getN <- function(x) sum(getUniques(x))
if (sample.cnt == 1){
  track <- cbind(out, getN(dadaFs), getN(dadaRs), getN(mergers), rowSums(seqtab.nochim))
  colnames(track) <- c("input", "filtered", "denoisedF", "denoisedR", "merged", "nonchim")
  rownames(track) <- sample.names
  head(track)
}else{
  track <- cbind(out, sapply(dadaFs, getN), sapply(dadaRs, getN), sapply(mergers, getN), rowSums(seqtab.nochim))
  colnames(track) <- c("input", "filtered", "denoisedF", "denoisedR", "merged", "nonchim")
  rownames(track) <- sample.names
  head(track)
}

# Store Step change File into csv
trackfile_name <- paste0(inpath, "/dadatrack_", format(Sys.Date(), "%Y%m%d"), ".csv")
write.csv(track, file = trackfile_name, row.names = TRUE)

# Create output folder
outfolder_name <- "dada2"
dir.create(file.path(inpath, outfolder_name))
dada2path <- paste0(inpath, "/dada2/")

for (i in 1:nrow(seqtab.nochim)) {
  # 設定檔名為行名稱
  dada2name <- paste0(dada2path, row.names(seqtab.nochim)[i], "_dada2.fasta")
  dada2trackname <- paste0(dada2path, row.names(seqtab.nochim)[i], "_dada2track.csv")

  # 尋找非零元素所在的column index
  nonzero_col <- which(seqtab.nochim[i,] != 0)
  file_out <- file(dada2name, open = "a")

  # 如果有非零元素，就將對應的column名稱寫入fasta檔中
  if (length(nonzero_col) > 0) {
    cntc <- 0
    col_names <- colnames(seqtab.nochim)[nonzero_col]
    for (col_name in col_names) {
      cntc = cntc + 1
      writeLines(paste0('>ASV',cntc,'\n',col_name,'\n'), file_out)
    }
    cntr <- length(nonzero_col)
    readnumlist <- data.frame(ASV = numeric(cntr), Value = character(cntr))
    readnums <- seqtab.nochim[i,nonzero_col]
    # 將編號1到N和list中的所有值寫入data frame
    total_sum <- sum(unlist(readnums))
    ratio <- readnums / total_sum
    readnumlist$ASV <- paste("ASV", 1:cntr, sep="")
    readnumlist$Value <- unlist(readnums)
    readnumlist$Ratio <- unlist(ratio)
    # 將data frame寫入csv檔案中
    write.csv(readnumlist, file = dada2trackname, row.names = FALSE)
  }
  close(file_out)
}
