from pytube import YouTube



class YoutubeDownloader():
    def __init__(self,video_url,out_dir) -> None:
        self.video_url = video_url
        self.out_dir = out_dir
        self.yt = YouTube(self.video_url,
                          use_oauth=True,
                          allow_oauth_cache=True).bypass_age_gate()
        self.stream = self.yt.streams.get_by_resolution("720p")

    def __enter__(self):
        try:
            self.stream.download(output_path=self.out_dir)
            print("Video download completed.\n")
            return f"{self.out_dir}/{self.yt.title}.mp4"
        except:
            print("An error has occured\n")
            exit()

    def __exit__(self,exc_type,exc_val,tracebac):
        print("Finished ...")

def execute():
    with YoutubeDownloader("https://www.youtube.com/watch?v=UKftOH54iNU&ab_channel=StromaeVEVO","C:/Users/LEGION/Desktop/Randomideas") as vd:
        print(vd)

if __name__ == "__main__":
    execute()