import docx2txt
import os

#Creating the image destination path
path = "Images"

# Check whether the specified path exists or not
isExist = os.path.exists(path)
if not isExist:

   # Create a new directory if it does not exist
   os.makedirs(path)


#take chapter number input
ChpNo=input('Chapter Number:')

#Take location of word file
wrdFile=input('Location of the file you want to process:\n')

#Take the filename to be extracted

file=input('Filename:')

#Combining the inputs to fetch the document

WordFinal = wrdFile + "\\" + file + ".docx"

#docx2txt Image extract program

text =  docx2txt.process(WordFinal, "Images")

#Renumberingthefiles

def rename():
    for count, filename in enumerate(os.listdir(path)):
        dst = f"Figure_{ChpNo}_{str(count+1)}.png"
        src =f"{path}/{filename}"  # foldername/filename, if .py file is outside folder
        dst =f"{path}/{dst}"
         
        # rename() function will
        # rename all the files
        os.rename(src, dst)
        print(dst + ' has been extracted')

# Driver Code
if __name__ == '__main__':
     
    # Calling rename() function
    rename()

