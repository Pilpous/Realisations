from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse
#from .models import ALBUMS
from .models import Album, Artist, Contact, Booking
from django.template import loader
from django.core.paginator import Paginator, PageNotAnInteger, EmptyPage
from .forms import ContactForm

# Create your views here.

def index(request):
    albums = Album.objects.filter(available=True).order_by('-created_at')[:12]
    formated_albums =  ["<li>{}</li>".format(album.title) for album in albums]
    #message = """<ul>{}</ul>""".format("\n".join(formated_albums))
    template = loader.get_template('store/index.html')
    context = {
        'albums': albums
    }
    #return HttpResponse(template.render(context, request=request)) 
    return render(request, 'store/index.html', context)

def listing(request):
    albums_list = Album.objects.filter(available=True)
    paginator = Paginator(albums_list, 1)
    page = request.GET.get('page')

    try:
        albums = paginator.page(page)
    except PageNotAnInteger:
        albums = paginator.page(1)
    except EmptyPage:
        albums = paginator.page(paginator.num_pages)
    #formated_albums =  ["<li>{}</li>".format(album.title) for album in albums]
    #message = """<ul>{}</ul>""".format("\n".join(formated_albums))
    context = {
        'albums': albums,
        'paginate' : True
    }
    #return HttpResponse(message)
    return render(request, 'store/listing.html', context)

def OLDdetail(request, album_id):
    album = get_object_or_404(Album, pk=album_id)
    artists = [artist.name for artist in album.artists.all()]
    artists_name = " ".join(artists)
    context = {
        'album_title': album.title,
        'artists_name': artists_name,
        'album_id': album.id,
        'thumbnail': album.picture
    }
    if request.method == 'POST':
        form = ContactForm(request.POST)
        if form.is_valid():
            email = form.cleaned_data['email']
            name = form.cleaned_data['name']

            contact = Contact.objects.filter(email=email)
            if not contact.exists():
                # If a contact is not registered, create a new one.
                contact = Contact.objects.create(
                    email=email,
                    name=name
                )
            else:
                contact = contact.first()

            album = get_object_or_404(Album, id=album_id)
            booking = Booking.objects.create(
                contact=contact,
                album=album
            )
            album.available = False
            album.save()
            context = {
                'album_title': album.title
            }
            return render(request, 'store/merci.html', context)
        else:
            # Form data doesn't match the expected format.
            # Add errors to the template.
            context['errors'] = form.errors.items()
    else:
        form = ContactForm()
    context['form'] = form
    return render(request, 'store/detail.html', context)

def detail(request, album_id):
    album = get_object_or_404(Album, pk=album_id)
    artists = [artist.name for artist in album.artists.all()]
    artists_name = " ".join(artists)

    if request.method == 'POST':
        email = request.POST.get('email')
        name = request.POST.get('name')
        contact = Contact.objects.filter(email=email)
        if not contact.exists():
            contact = Contact.objects.create(
                email=email,
                name=name
            )
        album = get_object_or_404(Album, id=album_id)
        booking = Booking.objects.create(
            contact=contact,
            album=album
        )
        album.available = False
        album.save()
        context = {
            'album_title' : album.title
        }
        return render(request, 'store/merci.html', context)
        #message = "Le nom de l'album est {}. Il a été écrit par {}".format(album.title, artists)
        #return HttpResponse(message)

    else:
        form = ContactForm()
    context = {
        'album_title': album.title,
        'artists_name': artists_name,
        'album_id': album.id,
        'thumbnail': album.picture,
        'form' : form
        }
    return render(request, 'store/detail.html', context)

def search(request):
    query = request.GET.get('query')
    if not query:
        albums = Album.objects.all()
    else :
        albums = Album.objects.filter(title__icontains=query)

    if not albums.exists():
        albums = Album.objects.filter(artists__name__icontains=query)

    title = "Résultats pour la requête %s"%query
    context = {
        'albums': albums,
        'title': title
    }
    #return HttpResponse(message)
    return render(request, 'store/search.html', context)