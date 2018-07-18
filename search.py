import sys
# append the possible next_state of parent_state to the list "children" with the parent
# i.e each element is a tuple of (state, parent of state)
def appendChildren(parent,children,x):
    if parent[x] == "9":
        newstr = parent[:x] + str(int(parent[x:x+1])-1) + parent[x+1:]
        children.append((newstr,parent,x))
    elif parent[x] == "0":
        newstr = parent[:x] + str(int(parent[x:x+1])+1) + parent[x+1:]
        children.append((newstr,parent,x))
    else:
        newstrPlus = parent[:x] + str(int(parent[x:x+1])+1) + parent[x+1:]
        newstrMinus = parent[:x] + str(int(parent[x:x+1])-1) + parent[x+1:]
        children.append((newstrMinus,parent,x))
        children.append((newstrPlus,parent,x))

def exploreChildren(parent,prevPos = None):
    children = []
    if prevPos is None:
        for x in range(3):
            appendChildren(parent,children,x)
    else:
        for index in list(filter(lambda x: x != prevPos, range(3))):
            appendChildren(parent,children,index)
    return  children

def BFS(start,goal,forbidden):
    # visited stores tuple of (state, parent of state)
    visited = [(start,None)]
    # toExplore stores tuple of (state, parent of state, previous index that's been changed)
    toExplore = [(start,None,None)]
    # for print
    exploredForPrint = []
    while (len(toExplore)>0):
        #print("toExplore: ",toExplore)
        current = toExplore.pop(0)
        exploredForPrint.append(current)
        #print("current: ",current)
        currentState,currentParent,prevPos = current[0],current[1],current[2]
        visitedState,parents = zip(*visited)
        candidates = []
        if currentState not in visitedState:
            candidates.extend(exploreChildren(currentState,prevPos))
            visited.append((currentState,currentParent))
            if currentState == goal:
                break
        elif currentState is start:
            candidates.extend(exploreChildren(currentState))
        toExplore.extend(list(filter(lambda x : x[0] not in forbidden, candidates)))

    shortestPath = []
    reverseCurrent = goal
    while(reverseCurrent != start):
        shortestPath.append(reverseCurrent)
        for x in range(len(visited)):
            if(visited[x][0] == reverseCurrent):
                reverseCurrent = visited[x][1]
                break
    shortestPath.append(start)
    shortestPath.reverse()
    print(shortestPath)
    print([item[0] for item in exploredForPrint])

def DFS(start,goal,forbidden):
    # visited stores tuple of (state, parent of state)
    visited = [(start,None)]
    # toExplore stores tuple of (state, parent of state, previous index that's been changed)
    toExplore = [(start,None,None)]
    # for print
    exploredForPrint = []
    while (len(toExplore)>0):
        #print("toExplore: ",toExplore)
        current = toExplore.pop(0)
        exploredForPrint.append(current)
        #print("current: ",current)
        currentState,currentParent,prevPos = current[0],current[1],current[2]
        visitedState,parents = zip(*visited)
        candidates = []
        if currentState not in visitedState:
            candidates.extend(exploreChildren(currentState,prevPos))
            visited.append((currentState,currentParent))
            if currentState == goal:
                break
        elif currentState is start:
            candidates.extend(exploreChildren(currentState))
        a = list(filter(lambda x : x[0] not in forbidden, candidates))
        toExplore = a + toExplore

    shortestPath = []
    reverseCurrent = goal
    while(reverseCurrent != start):
        shortestPath.append(reverseCurrent)
        for x in range(len(visited)):
            if(visited[x][0] == reverseCurrent):
                reverseCurrent = visited[x][1]
                break
    shortestPath.append(start)
    shortestPath.reverse()
    print(shortestPath)
    print([item[0] for item in exploredForPrint])

def DLS(start,goal,forbidden,depth):
    count = 0
    visited = [(start,None)]
    # toExplore stores tuple of (state, parent of state, previous index that's been changed)
    toExplore = [(start,None,None)]
    exploredForPrint = []
    foundgoal = False
    if depth == 0:
        return {"found":foundgoal,"visited":visited, "explored":toExplore}
    dep = 0
    while (len(toExplore)>0):
        current = toExplore.pop(0)
        if current == "flag":
            dep -= 1
            continue
        exploredForPrint.append(current)
        currentState,currentParent,prevPos = current[0],current[1],current[2]
        visitedState,parents = zip(*visited)
        candidates = []
        if currentState == goal:
            foundgoal = True
            visited.append((currentState,currentParent))
            break
        if dep < depth:
            if currentState not in visitedState:
                candidates.extend(exploreChildren(currentState,prevPos))
                dep+=1
                candidates.append("flag")
                visited.append((currentState,currentParent))
                if currentState == goal:
                    foundgoal = True
                    break
            elif currentState is start:
                candidates.extend(exploreChildren(currentState))
                dep+=1
            a = list(filter(lambda x : x[0] not in forbidden, candidates))
            toExplore = a + toExplore
    return {"found":foundgoal,"visited":visited, "explored":exploredForPrint}

def IDS(start,goal,forbidden):
    depth = 1
    resultOfDLS = []
    exploredForPrint = []
    for x in range(0,depth):
        resultOfDLS.append(DLS(start,goal,forbidden,x))
        exploredForPrint.extend(resultOfDLS[x]["explored"])
        if resultOfDLS[x]["found"]:
            break
    shortestPath = []
    reverseCurrent = goal
    visited = resultOfDLS[-1]["visited"]
    while(reverseCurrent != start):
        shortestPath.append(reverseCurrent)
        for x in range(len(visited)):
            if(visited[x][0] == reverseCurrent):
                reverseCurrent = visited[x][1]
                break
    shortestPath.append(start)
    shortestPath.reverse()
    print(shortestPath)
    print([item[0] for item in exploredForPrint])

def readFile(filename):
    with open(filename,"r",encoding = "utf-8") as file:
        lines = file.read().splitlines()
        backup = lines[2].split(",") if len(lines) == 3 else []
        return lines[0],lines[1],backup

def main():
    search_method, inputfile = sys.argv[1], sys.argv[2]
    start_state, goal_state, forbidden_states = readFile(inputfile)
    if(search_method == "B"):
        BFS(start_state,goal_state,forbidden_states)
    elif(search_method == "D"):
        DFS(start_state,goal_state,forbidden_states)
    elif search_method == "I":
        IDS(start_state,goal_state,forbidden_states)

if __name__ == "__main__":
    main()
