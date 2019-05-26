from PIL import Image, ImageDraw, ImageFont, ImageChops
import cv2
import numpy as np
import os

class Photo():

    def text_wrap(text, font, max_width):
        """
        if the text is to long to fit on 1 line, make more text lines
        :param text: text that is checked if it is too long
        :param font: text font
        :param max_width: maximum width of the text
        :return: list of lines with the text
        source: https://haptik.ai/tech/putting-text-on-images-using-python%E2%80%8A-%E2%80%8Apart2/
        """
        lines = []
        # If the width of the text is smaller than image width
        # we don't need to split it, just add it to the lines array
        # and return
        if font.getsize(text)[0] <= max_width:
            lines.append(text)
        else:
            text = text.replace('_', ' ')
            text = text.replace('-', ' ')
            # split the line by spaces to get words
            words = text.split(' ')
            i = 0
            # append every word to a line while its width is shorter than image width
            while i < len(words):
                line = ''
                while i < len(words) and font.getsize(line + words[i])[0] <= max_width:
                    line = line + words[i] + " "
                    i += 1
                if not line:
                    line = words[i]
                    i += 1
                # when the line gets longer than the max width do not append the word,
                # add the line to the lines array
                lines.append(line)
        return lines

    def makeTitle(title):
        """
        make a blue-gray image with a white title and save it with the name "title paramter".jpg
        :param title: name for the saved image + .jpg
        """
        # MAX_W, MAX_H = 6000, 4000
        MAX_W, MAX_H = 300, 300
        img = Image.new('RGB', (MAX_W, MAX_H), color = (73,109,137))
        d = ImageDraw.Draw(img)
        fnt = ImageFont.truetype('C:/Windows/Fonts/Arial.ttf', size=45, encoding='unic')
        color = 'rgb(255, 255, 255)'

        lines = Photo.text_wrap(title, fnt, MAX_W)

        linenr = 1
        for line in lines:
            w, h = fnt.getsize(line)
            x = (MAX_W - w) / 2
            y = (MAX_H - h) * linenr / (2*len(lines))
            # draw the line on the image
            d.text((x, y), line, fill=color, font=fnt, align='center')
            # update the y position so that we can use it for next line
            y = y + h
            linenr = linenr + 1
        # save the image
        img.save(title + '.jpg')
        # Photo.resize(title + '.jpg', title)

    def resize(img, title):
        """
        resize the image to have a basewidth of 300 and the height resized to keep the ratio.
        save it with the name "title parameter".jpg
        source: https://stackoverflow.com/questions/273946/how-do-i-resize-an-image-using-pil-and-maintain-its-aspect-ratio
        :param img: imgage that is being resized
        :param title: name for the saved image + .jpg
        """
        print('image')
        print(img)

        img = Photo.crop(img)
        # img = Image.open(img)
        # basewidth = 300
        # wpercent = (basewidth / float(img.size[0]))
        # hsize = int((float(img.size[1]) * float(wpercent)))
        basewidth = 300
        hsize = 300
        img = img.resize((basewidth, hsize), Image.ANTIALIAS)
        img.save(title + '.jpg')


    def merge_image(img1, img2, vertically, name):
        """
        merge_image takes three parameters first two parameters specify
        the two images to be merged and third parameter i.e. vertically
        is a boolean type which if True merges images vertically
        and finally saves and returns the file_name
        source: https://stackoverflow.com/questions/30227466/combine-several-images-horizontally-with-python
        :param img1: first image that is being merged
        :param img2: second image that is being merged
        :param vertically: boolean. True: image 2 is merged under image 1. False: image 2 is merged to the right of image 1
        :param name: name for the saved image + .jpg
        """
        images = list(map(Image.open, [img1, img2]))
        widths, heights = zip(*(i.size for i in images))
        if vertically:
            max_width = max(widths)
            total_height = sum(heights)
            new_im = Image.new('RGB', (max_width, total_height))

            y_offset = 0
            for im in images:
                new_im.paste(im, (0, y_offset))
                y_offset += im.size[1]
        else:
            total_width = sum(widths)
            max_height = max(heights)
            new_im = Image.new('RGB', (total_width, max_height))

            x_offset = 0
            for im in images:
                new_im.paste(im, (x_offset, 0))
                x_offset += im.size[0]

        new_im.save(name + '.jpg')


    def crop(image):
        print(image)
        img = cv2.imread(image)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        cv2.imwrite('gray.jpg', gray)
        imgG = Image.open('gray.jpg')
        r = np.array(imgG)
        w, h = imgG.size
        print(w)
        print(h)
        # zero_row_indices = [i for i in r.shape[0] if np.allclose(r[i,:],0)]
        # nonzero_row_indices =[i for i in r.shape[0] if not np.allclose(r[i,:],0)]
        print(r.shape)
        print(np.mean(r))
        limit = np.mean(r)
        step = 100
        procent = 20
        col1 = 1 - step
        col2 = w-1 + step #6000 + step
        firstCol = False
        for y in range(1, w-1, step):
            col_vec = []
            for x in range(1, h-1):
                if r[x, y] < limit:
                    col_vec.append(x)

            if len(col_vec) > h-h/procent:
                # print(y)
                # print(col1)
                # print(col2)
                # print('---')
                if col1 == y - step:
                    col1 = y
                else:
                    firstCol = True
            if firstCol:
                break

        secondCol = False
        for y in range(w-1, 1, -step):
            col_vec = []
            for x in range(1, h-1):
                if r[x, y] < limit:
                    col_vec.append(x)
            if len(col_vec) > h-h/procent:
                # print(y)
                # print(col1)
                # print(col2)
                # print('---')
                if col2 == y + step:
                    col2 = y
                else:
                    secondCol = True
            if secondCol:
                break

        row1 = 1 - step
        row2 = h-1 + step #4000 + step
        firstRow = False
        for x in range(1, h-1, step):
            row_vec = []
            for y in range(1, w-1):
                if r[x, y] < limit:
                    row_vec.append(x)

            if len(row_vec) > w-w/procent:
                # print(y)
                # print(col1)
                # print(col2)
                # print('---')
                if row1 == x - step:
                    row1 = x
                else:
                    firstRow = True
            if firstRow:
                break

        secondRow = False
        for x in range(h-1, 1, -step):
            row_vec = []
            for y in range(1, w-1):
                if r[x, y] < limit:
                    row_vec.append(x)
            if len(row_vec) > w-w/procent:
                # print(y)
                # print(col1)
                # print(col2)
                # print('---')
                if row2 == x + step:
                    row2 = x
                else:
                    secondRow = True
            if secondRow:
                break
        os.remove('gray.jpg')
        img = Image.open(image)
        img = img.crop((col1, row1, col2, row2))
        return img