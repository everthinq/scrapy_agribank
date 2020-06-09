from scrapy.cmdline import execute

def main():
    # execute(['scrapy', 'crawl', 'agribank', '--nolog'])
    execute(['scrapy', 'crawl', 'agribank'])

if __name__ == '__main__':
    main()