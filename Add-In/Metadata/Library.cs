using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMPLib;
using VmcController.AddIn.Commands;
using VmcController.AddIn.Metadata;


namespace VmcController.AddIn
{
    public class Library : OpResultObject
    {
        public bool from_cache = false;
        private bool stats_only = false;

        public ArrayList albums;
        public ArrayList songs;
        public ArrayList genres;
        public ArrayList artists;
        public ArrayList album_artists;
        public ArrayList playlists;


        public Library(string param)
        {
            if (param.Contains("stats")) stats_only = true;
            else stats_only = false;
        }

        public bool is_stats_only
        {
            set { stats_only = value; }
            get { return stats_only; }
        }

        public void addSongs(IWMPPlaylist playlist, ArrayList indexes)
        {
            songs = new ArrayList();
            for (int j = 0; j < indexes.Count; j++)
            {
                IWMPMedia item = playlist.get_Item((Int16)indexes[j]);
                if (item != null)
                {
                    songs.Add(new Song(item));
                }
            }
            songs.TrimToSize();
        }

        public int addSongs(IWMPPlaylist playlist)
        {
            ArrayList items = new ArrayList();
            for (int j = 0; j < playlist.count; j++)
            {
                IWMPMedia song = playlist.get_Item(j);
                if (song != null)
                {
                    Song item;
                    if (!stats_only)
                    {
                        item = new Song(song);                        
                    }
                    else
                    {
                        item = new Song(song.name);
                    }
                    if (song.name != null && !song.name.Equals("") && !items.Contains(item))
                    {
                        items.Add(item);
                    }
                }
            }
            items.TrimToSize();
            if (!stats_only)
            {
                songs = items;
            }
            return items.Count;
        }

        private int getArrayCount(ArrayList items)
        {
            return items.Count;
        }

        private ArrayList getArray(IWMPStringCollection collection)
        {
            ArrayList items = new ArrayList();
            for (int k = 0; k < collection.count; k++)
            {
                string item = collection.Item(k);
                if (item != null && !item.Equals("") && !items.Contains(item))
                {
                    items.Add(item);
                }
            }
            items.TrimToSize();
            return items;
        }

        public int addAlbums(IWMPStringCollection collection, IWMPMediaCollection2 mediaCollection)
        {
            ArrayList items = new ArrayList();
            for (int k = 0; k < collection.count; k++)
            {
                string name = collection.Item(k);
                if (name != null && !name.Equals(""))
                {
                    Album item;
                    if (!stats_only)
                    {
                        item = new Album(name, mediaCollection.getByAlbum(name));
                    }
                    else
                    {
                        item = new Album(name, true); 
                    }
                    if (!items.Contains(item))
                    {
                        items.Add(item); 
                    }
                }
            }
            items.TrimToSize();
            if (!stats_only)
            {
                albums = items;
            }
            return items.Count;
        }

        public int addGenres(IWMPStringCollection collection)
        {
            int result_count = 0;
            if (!stats_only)
            {
                genres = getArray(collection);
                result_count = genres.Count;
            }
            else result_count = getArrayCount(getArray(collection));
            return result_count;
        }

        public int addArtists(IWMPStringCollection collection)
        {
            int result_count = 0;
            if (!stats_only)
            {
                artists = getArray(collection);
                result_count = artists.Count;
            }
            else result_count = getArrayCount(getArray(collection));
            return result_count;
        }

        public int addAlbumArtists(IWMPStringCollection collection)
        {
            int result_count = 0;
            if (!stats_only)
            {
                album_artists = getArray(collection);
                result_count = album_artists.Count;
            }
            else result_count = getArrayCount(getArray(collection));
            return result_count;
        }

        public int addPlaylists(IWMPPlaylistCollection playlistCollection, IWMPPlaylistArray list)
        {
            ArrayList items = new ArrayList();
            for (int j = 0; j < list.count; j++)
            {
                bool containsAudio = false;
                IWMPPlaylist playlist = list.Item(j);
                string name = playlist.name;

                if (!name.Equals("All Music") && !name.Contains("TV") && !name.Contains("Video") && !name.Contains("Pictures"))
                {
                    for (int k = 0; k < playlist.count; k++)
                    {
                        try
                        {
                            if (playlist.get_Item(k).getItemInfo("MediaType").Equals("audio") && !playlistCollection.isDeleted(playlist))
                            {
                                containsAudio = true;
                            }
                        }
                        catch (Exception)
                        {
                            //Ignore playlists with invalid items
                        }
                    }
                }

                if (containsAudio)
                {
                    Playlist playlistData = new Playlist(name, stats_only);
                    if (!items.Contains(playlistData)) items.Add(playlistData);                    
                }
            }
            items.TrimToSize();
            if (!stats_only)
            {
                playlists = items;
            }
            return items.Count;
        }
    }
}
