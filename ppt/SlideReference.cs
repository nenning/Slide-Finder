namespace ppt
{
    class SlideReference
    {
        public static float MaxImageDifference = 0.65f;
        public static int MaxTextDistance = 100;
        public SlideReference(int sourceSlide)
        {
            SourceSlide = sourceSlide;
            TextSlideRef = -1;
            ImageSlideRef = -1;
            TextDistance = MaxTextDistance;
            ImageDifference = 1f;
        }

        public int SourceSlide { get; private set; }
        public int TextSlideRef { get; private set; }
        public int ImageSlideRef { get; private set; }
        public int TextDistance { get; private set; }
        public float ImageDifference { get; private set; }

        public void UpdateImage(int slideRef, float difference)
        {
            if (difference < MaxImageDifference && difference < ImageDifference)
            {
                ImageDifference = difference;
                ImageSlideRef = slideRef;
            }
        }
        public void UpdateText(int slideRef, int distance)
        {
            // Do a text size comparison(maybe also match)?
            // store 2 matches each?
            if (distance < MaxTextDistance && distance < TextDistance)
            {
                TextDistance = distance;
                TextSlideRef = slideRef;
            }
        }
        public override string ToString()
        {
            int percentage = (int)(ImageDifference * 100);

            if (TextSlideRef == -1 && ImageSlideRef == -1)
            {
                return $"slide #{SourceSlide,3} -> n/a";
            }
            else
            {
                int slide1 = ImageSlideRef;
                int slide2 = -1;
                if (slide1 == -1) slide1 = TextSlideRef;
                else slide2 = TextSlideRef;
                if (slide1 == slide2 || slide2 == -1)
                {
                    return $"slide #{SourceSlide,3} -> try master slide #{GetSlideRef(slide1)}              (Diff: {percentage}%, Dist: {TextDistance})";
                }
                else
                {
                    return $"slide #{SourceSlide,3} -> try master slide #{GetSlideRef(ImageSlideRef)} or #{GetSlideRef(TextSlideRef)}      (Diff: {percentage}%, Dist: {TextDistance})";
                }
            }
        }

        private string GetSlideRef(int slide)
        {
            if (slide <= 0) return "n/a";
            else return $"{slide,3}";
        }
    }
}
