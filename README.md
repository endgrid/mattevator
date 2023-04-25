# mattevator
![elevator](https://user-images.githubusercontent.com/104172903/234367598-3d099e89-92dc-4711-bca8-bd831b90c3cf.png)

Consider a busy two story building with a single elevator. In the morning, some will use the elevator to get upstairs.

```
___________________
|   |   MATTCORP  |
|   |  -->        |
|_@_|_______@@@@@_|
| ^ |             |
| ^ |  <--        |
|_@_|______@@@@@__|  <-- @@@@@
```

In the afternoon, a smaller number will use the elevator to get downstairs (some will take stairs).

```
___________________
|   |   MATTCORP  |
|   |  <--        |
|_@_|_______@@@___|
| v |             |
| v |  -->        |
|_@_|______@@@@@__|  --> @@@@@
```

A facilities maintenance employee notices that after an elevator successfully delivers a passenger, it returns to the ground floor.

```
___________________
|   |   MATTCORP  |
|   |             |
|___|_____________|
| X |             |
| X |             |
|_X_|_____________|
```

In the morning, the ground floor is the optimal elevator resting position, as those entering the building do not have to wait for the elevator to arrive.

```
___________________
|   |   MATTCORP  |
|   |             |
|___|_____________|
| X |             |
| X |  nice       |
|_X_|_@___________|
```

In the afternoon, the employee notices, no one is arriving at the building, and those attempting to exit the second floor via elevator experience a delay.

```
___________________
|   |   MATTCORP  |
|   |  ?          |
|___|_@___________|
| ^ |             |
| ^ |             |
|_^_|_____________|
```

They have to wait for the elevator to arrive before descending in it, compared to their journey in the morning where the elevator was waiting for them.

```
___________________
| x |   MATTCORP  |
| x |  nice       |
|_x_|_@___________|
|   |             |
|   |             |
|___|_____________|
```

The employee reprograms the elevator to return to the second floor instead of the first upon the completion of a successful passenger delivery in the afternoon.

```
___________________
| x |   MATTCORP  |
| x |             |
|_x_|_____________|
|   |             |
|   |             |
|___|_____________|
```

Now, no one has to wait for an elevator that is not in use to arrive, as its resting positions have been optimized.

```
___________________
| x |   MATTCORP  |
| x |  nice       |
|_x_|_@@@_________|
|   |             |
|   |             |
|___|_____________|
```

The is the crux of mattevator--Given the number of travellers predicted to leave a floor via elevator (Referred to as "emigrators" in the code) over a given period of time, we can create an elevator resting position optimization model that minimizes travel time by mitigating the delay periods of unoccupied elevator arrival.

## Walkthrough:
<img width="641" alt="image" src="https://user-images.githubusercontent.com/104172903/234387290-88d74498-995c-497e-a7ad-ff112c085e17.png">

![image](https://user-images.githubusercontent.com/104172903/234386004-9b291f5f-eb82-4bf3-ac30-05bf5e6006ac.png)

MATTEVATOR runs as a macro enabled excel workbook.
The landing page, the MATTEVATOR HOME tab, gives you the ability to select the number of elevators and floors in the building.

Clicking the "First time" button takes you to the MATTEVATOR LEARN tab, which details the function of the application.

![image](https://user-images.githubusercontent.com/104172903/234387400-0e23d0cd-1eb8-49f7-b328-f3c553026a72.png)

Clicking "Optimize" prompts the user to enter the number of emigrators per floor:

![image](https://user-images.githubusercontent.com/104172903/234387819-c29d9c71-2a77-43d0-82d9-007658d5c8cf.png)

In this case, I entered 5 emigrating from floor one, and zero emigrating from floor two, to simulate the morning scenario above.

Mattevator then calculates the optimal resting position, and places an X where it thinks the elevator should rest.

![image](https://user-images.githubusercontent.com/104172903/234387920-347b2446-5ddc-47cb-b8e9-c6b129db0ca3.png)

It also pulls from the New York Elevator Database to identify the building in New York that is most similar to yours based on number of elevators.

![image](https://user-images.githubusercontent.com/104172903/234396256-f1ebde81-e022-43ca-89c9-186cc393581e.png)

Mattevator is capable of calculating optimal elevator resting positions for very large buildings, such as one with a 75 floors and eight elevators:

![image](https://user-images.githubusercontent.com/104172903/234396905-430a7e16-c858-451f-81b6-2a72ab5971c4.png)

![image](https://user-images.githubusercontent.com/104172903/234397823-429f8824-519c-4244-904d-4c6526987414.png)


![image](https://user-images.githubusercontent.com/104172903/234397719-9ff95a5d-26cd-493b-b08f-8e5e2bd5bc23.png)





