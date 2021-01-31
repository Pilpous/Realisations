from django import forms

class ContactForm(forms.Form):
    name = forms.CharField(
        label='Nom',
        max_length=100,
        widget=forms.TextInput(attrs={'class' : 'forms-control'}),
        required=True
        )
    email = forms.EmailField(
        label='Email',
        max_length=100,
        widget=forms.EmailInput(attrs={'class' : 'forms-control'}),
        required=True
    )