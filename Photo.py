from PIL import Image, ImageDraw, ImageFont, ImageChops
import cv2
import numpy as np
import os
import imutils

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
        try:
            img = Photo.transform(img)
        except:
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
        # cv2.imwrite('gray.jpg', gray)
        # imgG = Image.open('gray.jpg')
        r = np.array(gray)
        w = r.shape[1]
        h = r.shape[0]
        print(w)
        print(h)
        # zero_row_indices = [i for i in r.shape[0] if np.allclose(r[i,:],0)]
        # nonzero_row_indices =[i for i in r.shape[0] if not np.allclose(r[i,:],0)]
        print(r.shape)
        print(np.mean(r))
        threshold = np.mean(r)
        step = 100
        procent = 20
        col1 = 1 - step
        col2 = w-1 + step #6000 + step
        firstCol = False
        for y in range(1, w-1, step):
            col_vec = []
            for x in range(1, h-1):
                if r[x, y] < threshold:
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
                if r[x, y] < threshold:
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
                if r[x, y] < threshold:
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
                if r[x, y] < threshold:
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

        # os.remove('gray.jpg')

        # img = Image.open(image)
        img = Image.fromarray(img)
        img = img.crop((col1, row1, col2, row2))

        return img

    def transform(image):
        """
        source: https://www.pyimagesearch.com/2014/09/01/build-kick-ass-mobile-document-scanner-just-5-minutes/
        :return:
        """
        image = cv2.imread(image)

        ratio = image.shape[0] / 500.0
        orig = image.copy()
        image = imutils.resize(image, height=500)

        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

        r = np.array(gray)
        print(np.mean(r))
        threshold = np.mean(r)

        ret, thresh = cv2.threshold(gray, threshold, 255, cv2.THRESH_BINARY)


        # try:
        # small closing
        kernel = np.ones((15, 15), np.uint8)
        closing = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel, borderValue=0)

        # large opening
        kernel = np.ones((25, 25), np.uint8)
        opening = cv2.morphologyEx(closing, cv2.MORPH_OPEN, kernel, borderValue=0)

        # large closing
        kernel = np.ones((35, 35), np.uint8)
        closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel, borderValue=0)

        # convert the image to grayscale, blur it, and find edges
        # in the image
        # gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(closing, (9, 9), cv2.BORDER_DEFAULT)
        edged = cv2.Canny(gray, 0, 20)
        edged = cv2.GaussianBlur(edged, (9, 9), cv2.BORDER_DEFAULT)

        (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        total = 0
        for c in cnts:
            epsilon = 0.08 * cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, epsilon, True)

            cv2.drawContours(image, [approx], -1, (0, 255, 0), 4)
            total += 1
        screenCnt = approx
        # (cnts, _) = cv2.findContours(edged.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        #
        #
        # total = 0
        # for c in cnts:
        #     epsilon = 0.08 * cv2.arcLength(c, True)
        #     approx = cv2.approxPolyDP(c, epsilon, True)
        #
        #     cv2.drawContours(image, [approx], -1, (0, 255, 0), 4)
        #     total += 1
        #
        # print("I found {0} RET in that image".format(total))
        # cv2.imshow("Output", image)
        # cv2.waitKey(0)
        # exit()

        # show the original image and the edge detected image
        print("STEP 1: Edge Detection")
        # cv2.imshow("blur", gray)
        # cv2.imshow("Image", image)
        # cv2.imshow("Edged", edged)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        # -------------------------------------------------

        # find the contours in the edged image, keeping only the
        # largest ones, and initialize the screen contour
        # cnts = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
        # cnts = imutils.grab_contours(cnts)
        # cnts = sorted(cnts, key=cv2.contourArea, reverse=True)[:5]
        #
        # # loop over the contours
        # for c in cnts:
        #     # approximate the contour
        #     peri = cv2.arcLength(c, True)
        #     epsilon = 0.02*peri
        #     approx = cv2.approxPolyDP(c, epsilon * peri, True)
        #
        #     # if our approximated contour has four points, then we
        #     # can assume that we have found our screen
        #     if len(approx) == 4:
        #         screenCnt = approx
        #         break
        # return Photo.__transformFunc(orig, screenCnt, ratio)

        # show the contour (outline) of the piece of paper
        print("STEP 2: Find contours of paper")
        # cv2.drawContours(image, [screenCnt], -1, (0, 255, 0), 2)
        # cv2.imshow("Outline", image)
        # cv2.waitKey(0)
        # cv2.destroyAllWindows()

        # apply the four point transform to obtain a top-down
        # view of the original image
        # warped = four_point_transform(orig, screenCnt.reshape(4, 2) * ratio) #-------------------------------------------
        image = orig
        pts = screenCnt.reshape(4, 2) * ratio
        # obtain a consistent order of the points and unpack them
        # individually
        # initialzie a list of coordinates that will be ordered
        # such that the first entry in the list is the top-left,
        # the second entry is the top-right, the third is the
        # bottom-right, and the fourth is the bottom-left
        rect = np.zeros((4, 2), dtype="float32")

        # the top-left point will have the smallest sum, whereas
        # the bottom-right point will have the largest sum
        s = pts.sum(axis=1)
        rect[0] = pts[np.argmin(s)]
        rect[2] = pts[np.argmax(s)]

        # now, compute the difference between the points, the
        # top-right point will have the smallest difference,
        # whereas the bottom-left will have the largest difference
        diff = np.diff(pts, axis=1)
        rect[1] = pts[np.argmin(diff)]
        rect[3] = pts[np.argmax(diff)]

        # return the ordered coordinates
        # rect = order_points(pts) #-------------------------------
        (tl, tr, br, bl) = rect

        # compute the width of the new image, which will be the
        # maximum distance between bottom-right and bottom-left
        # x-coordiates or the top-right and top-left x-coordinates
        widthA = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
        widthB = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
        maxWidth = max(int(widthA), int(widthB))

        # compute the height of the new image, which will be the
        # maximum distance between the top-right and bottom-right
        # y-coordinates or the top-left and bottom-left y-coordinates
        heightA = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
        heightB = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
        maxHeight = max(int(heightA), int(heightB))

        # now that we have the dimensions of the new image, construct
        # the set of destination points to obtain a "birds eye view",
        # (i.e. top-down view) of the image, again specifying points
        # in the top-left, top-right, bottom-right, and bottom-left
        # order
        dst = np.array([
            [0, 0],
            [maxWidth - 1, 0],
            [maxWidth - 1, maxHeight - 1],
            [0, maxHeight - 1]], dtype="float32")

        # compute the perspective transform matrix and then apply it
        M = cv2.getPerspectiveTransform(rect, dst)
        warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
        warped = Image.fromarray(warped)
        return warped
