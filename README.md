## Dizemanepe App




find . -type d -name '__pycache__' -exec rm -r {} +
find . -type d -name 'migrations' -not -path "./.venv/*" -exec rm -r {} +
rm mydatabase

python3 manage.py makemigrations
python3 manage.py makemigrations app
python3 manage.py makemigrations users
python3 manage.py makemigrations costumes
python3 manage.py migrate
python3 manage.py create_superuser alex asdfghj1!
python3 manage.py create_superuser iulian asdfghj1!
python3 manage.py create_groups
python3 manage.py create_colors
python3 manage.py create_test_clients
python3 manage.py create_event_types
python3 manage.py create_costume_states
python3 manage.py create_costumes
python3 manage.py create_services






















https://dizemanepe-photoshare.s3.eu-north-1.amazonaws.com/userAvatar.png


<select name="color" id="{{costumeStateForm.color.id_for_label}}">
    {% for value, display in costumeStateForm.color.field.choices %}
    <option value="{{ value }}" class="{{display}} ">{{ display }}</option>
    {% endfor %}
</select>
