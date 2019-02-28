from PIL import Image, ImageDraw, ImageFont

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
        MAX_W, MAX_H = 6000, 4000
        img = Image.new('RGB', (MAX_W, MAX_H), color = (73,109,137))
        d = ImageDraw.Draw(img)
        fnt = ImageFont.truetype('C:/Windows/Fonts/Arial.ttf', size=750, encoding='unic')
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
        Photo.resize(title + '.jpg', title)

    def resize(img, title):
        """
        resize the image to have a basewidth of 300 and the height resized to keep the ratio.
        save it with the name "title parameter".jpg
        source: https://stackoverflow.com/questions/273946/how-do-i-resize-an-image-using-pil-and-maintain-its-aspect-ratio
        :param img: imgage that is being resized
        :param title: name for the saved image + .jpg
        """
        img = Image.open(img)
        basewidth = 300
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
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